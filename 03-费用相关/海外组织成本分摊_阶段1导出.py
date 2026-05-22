import os
import time
import asyncio
from pathlib import Path
from getpass import getpass

from playwright.async_api import Playwright, async_playwright, TimeoutError as PlaywrightTimeoutError

BASE_DIR = Path(__file__).resolve().parent

# 配置路径
DOWNLOAD_DIR = Path(r"C:\Users\huangxiaoyan\Documents\trae_projects\test_week\downloads")
DOWNLOAD_DIR.mkdir(exist_ok=True)

# 报表配置
REPORT_CONFIG = {
    'name': '海外漏斗达成情况-月趋势',
    'url': 'https://bi.61info.cn/smartbi/vision/openresource.jsp?resid=I2c928087019e186e186e355a019e1f2b6636131d'
}

# 区域配置
REGION_CONFIGS = [
    {
        'name': '整体',
        'regions': ['北美', '欧洲', '海外其他', '亚洲', '澳洲', '国内', '港澳', '台湾']
    },
    {
        'name': '非台湾',
        'regions': ['北美', '欧洲', '海外其他', '亚洲', '澳洲', '国内', '港澳']
    },
    {
        'name': '欧美澳',
        'regions': ['北美', '欧洲', '海外其他', '亚洲', '澳洲', '国内']
    },
    {
        'name': '港澳',
        'regions': ['港澳']
    },
    {
        'name': '台湾',
        'regions': ['台湾']
    }
]


def load_config():
    """加载配置"""
    env_path = BASE_DIR / ".env"
    if env_path.exists():
        with open(env_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    if '=' in line:
                        key, value = line.split('=', 1)
                        os.environ[key.strip()] = value.strip()

    username = os.environ.get('SMARTBI_USERNAME')
    password = os.environ.get('SMARTBI_PASSWORD')
    base_url = os.environ.get('SMARTBI_BASE_URL', 'https://bi.61info.cn/smartbi/vision/index.jsp')

    if not username:
        username = input("请输入 Smartbi 用户名: ")
    if not password:
        password = getpass("请输入 Smartbi 密码: ")

    return {
        'username': username,
        'password': password,
        'base_url': base_url,
        'download_dir': DOWNLOAD_DIR
    }


async def login(page, config):
    """登录SmartBI"""
    print("\n" + "="*60)
    print("统一登录 Smartbi")
    print("="*60)
    
    await page.goto(config['base_url'], wait_until="networkidle")

    try:
        await page.wait_for_selector('input[bofid="username"]', timeout=15000)
    except PlaywrightTimeoutError:
        raise Exception("登录页面加载超时")

    await page.fill('input[bofid="username"]', config['username'])
    await page.fill('input[bofid="password"]', config['password'])
    await page.click('input[bofid="login"]')

    await page.wait_for_load_state("networkidle", timeout=30000)
    print("[SUCCESS] 登录成功")


async def select_month_filter(page, start_month, end_month):
    """选择月份筛选条件 - 使用直接输入方式"""
    print(f"[INFO] 设置月份筛选: {start_month} 至 {end_month}")
    
    try:
        # 查找日历输入框
        month_inputs = page.locator('input.multicalendar-nobackground-edit')
        count = await month_inputs.count()
        print(f"[DEBUG] 找到 {count} 个月份输入框")
        
        if count >= 2:
            # 设置开始月份
            start_input = month_inputs.nth(0)
            await start_input.evaluate('el => el.removeAttribute("readonly")')
            await start_input.fill(start_month)
            await start_input.dispatch_event('change')
            await start_input.dispatch_event('blur')
            print(f"[DEBUG] 已设置开始月份: {start_month}")
            
            # 设置结束月份
            end_input = month_inputs.nth(1)
            await end_input.evaluate('el => el.removeAttribute("readonly")')
            await end_input.fill(end_month)
            await end_input.dispatch_event('change')
            await end_input.dispatch_event('blur')
            print(f"[DEBUG] 已设置结束月份: {end_month}")
            
            await asyncio.sleep(1)
        else:
            print(f"[WARNING] 未找到足够的月份输入框，找到: {count} 个")
            
    except Exception as e:
        print(f"[ERROR] 设置月份筛选失败: {e}")


async def select_region_filter(page, regions):
    """选择区域等级筛选条件 - 直接输入文本（同外渠报表批量导出方式）"""
    print(f"[INFO] 设置区域筛选: {regions}")
    try:
        print("[DEBUG] 等待筛选控件加载...")
        await page.wait_for_selector('input.combobox-edit, .combobox-panel', timeout=30000)

        target_input = None

        # 方法1（优先）：通过“区域等级”标签行 + XPath 精确定位
        region_input_xpath = page.locator(
            'xpath=//td[contains(normalize-space(.),"区域等级")]/following-sibling::td//input[contains(@class,"combobox-edit")]'
        )
        xpath_count = await region_input_xpath.count()
        print(f"[DEBUG] 区域等级 XPath 命中输入框数量: {xpath_count}")
        for i in range(xpath_count):
            cand = region_input_xpath.nth(i)
            if await cand.is_visible():
                target_input = cand
                print(f"[DEBUG] 命中区域等级输入框(xpath) index={i}")
                break

        # 方法2：按“区域等级”所在行定位（兼容DOM结构变化）
        if not target_input or await target_input.count() == 0:
            region_label = page.locator('td:has-text("区域等级")').first
            if await region_label.count() > 0:
                tr = region_label.locator('xpath=ancestor::tr[1]')
                nearby = tr.locator('input.combobox-edit')
                near_count = await nearby.count()
                print(f"[DEBUG] 区域等级同行找到 {near_count} 个输入框")
                for i in range(near_count):
                    cand = nearby.nth(i)
                    if await cand.is_visible():
                        target_input = cand
                        print(f"[DEBUG] 命中区域等级输入框(tr) index={i}")
                        break

        # 方法3：兜底，仅取第一个可见输入框（避免直接超时退出）
        if not target_input or await target_input.count() == 0:
            visible_inputs = page.locator('input.combobox-edit:visible')
            vis_count = await visible_inputs.count()
            print(f"[WARN] 区域等级精确定位失败，兜底可见输入框数量: {vis_count}")
            if vis_count > 0:
                target_input = visible_inputs.first
                print("[WARN] 使用兜底输入框 index=0")

        if not target_input or await target_input.count() == 0:
            raise Exception("未找到区域等级输入框")

        print("[DEBUG] 定位到区域等级输入框")
        regions_text = ",".join(regions)
        await target_input.click(timeout=5000, force=True)
        try:
            await target_input.fill("")
        except Exception:
            await target_input.press("Control+a")
            await target_input.press("Backspace")

        await target_input.type(regions_text, delay=30)
        print(f"[DEBUG] 已输入区域文本: {regions_text}")

        await target_input.press("Enter")
        await page.click("body", position={"x": 50, "y": 50}, timeout=3000)
        await asyncio.sleep(1.5)

        current_value = (await target_input.input_value() or "").strip()
        print(f"[DEBUG] 输入框当前值: {current_value}")
        missing = [r for r in regions if r not in current_value]
        if missing:
            print(f"[WARN] 输入框未完整回显这些区域（可能是展示格式差异）: {missing}")

        # 防止“区域细分”串值：强制重置为“全部”
        try:
            detail_label = page.locator('td:has-text("区域细分")').first
            if await detail_label.count() > 0:
                detail_tr = detail_label.locator('xpath=ancestor::tr[1]')
                detail_inputs = detail_tr.locator('input.combobox-edit')
                detail_input = None
                detail_count = await detail_inputs.count()
                for i in range(detail_count):
                    cand = detail_inputs.nth(i)
                    if await cand.is_visible():
                        detail_input = cand
                        break
                if detail_input and await detail_input.count() > 0:
                    await detail_input.click(timeout=3000, force=True)
                    try:
                        await detail_input.fill("")
                    except Exception:
                        await detail_input.press("Control+a")
                        await detail_input.press("Backspace")
                    await detail_input.type("全部", delay=20)
                    await detail_input.press("Enter")
                    await page.click("body", position={"x": 60, "y": 60}, timeout=2000)
                    await asyncio.sleep(0.5)
                    print("[DEBUG] 已将区域细分重置为: 全部")
        except Exception as e:
            print(f"[WARN] 区域细分重置失败(继续): {e}")

        print("[DEBUG] 区域筛选完成")

    except PlaywrightTimeoutError as e:
        raise Exception(f"操作超时: {e}")
    except Exception as e:
        print(f"[ERROR] 设置区域筛选失败: {e}")
        raise


async def refresh_report(page):
    """刷新报表"""
    print("[INFO] 刷新报表")
    
    try:
        # 优先查找 SmartBI 特定的刷新按钮
        smartbi_refresh = page.locator('input.btnRefresh, button.btnRefresh, input.queryview-toolbar-refresh, button.queryview-toolbar-refresh')
        if await smartbi_refresh.count() > 0:
            await smartbi_refresh.first.click()
            print("[DEBUG] 点击 SmartBI 刷新按钮")
            # SmartBI 常有长连接，networkidle容易假超时
            try:
                await page.wait_for_load_state("domcontentloaded", timeout=10000)
            except Exception:
                pass
            await asyncio.sleep(8)
            return
        
        # 查找刷新按钮
        refresh_buttons = page.locator('button, input[type="button"], input[type="submit"]')
        count = await refresh_buttons.count()
        print(f"[DEBUG] 找到 {count} 个按钮")
        
        # 尝试点击刷新按钮
        for i in range(count):
            button = refresh_buttons.nth(i)
            try:
                text = await button.text_content()
                if text and ('刷新' in text or 'Refresh' in text):
                    await button.click()
                    print("[DEBUG] 点击刷新按钮")
                    try:
                        await page.wait_for_load_state("domcontentloaded", timeout=10000)
                    except Exception:
                        pass
                    await asyncio.sleep(8)
                    return
            except Exception as e:
                continue
        
        # 备用方案：查找图标按钮
        icon_buttons = page.locator('[title="刷新"], [title="Refresh"]')
        if await icon_buttons.count() > 0:
            await icon_buttons.first.click()
            print("[DEBUG] 点击图标刷新按钮")
            try:
                await page.wait_for_load_state("domcontentloaded", timeout=10000)
            except Exception:
                pass
            await asyncio.sleep(8)
            return
        
        print("[WARNING] 未找到刷新按钮，尝试按F5刷新")
        await page.keyboard.press('F5')
        try:
            await page.wait_for_load_state("domcontentloaded", timeout=15000)
        except Exception:
            pass
        await asyncio.sleep(8)
        
    except Exception as e:
        print(f"[ERROR] 刷新报表失败: {e}")


async def export_report(page, config, report_name, region_name) -> Path:
    """导出报表并按规则命名"""
    download_file_path = None
    download_event_handler = None

    async def handle_download(download):
        nonlocal download_file_path
        file_name = download.suggested_filename or f"report_{int(time.time())}.xlsx"
        date_suffix = time.strftime("%y%m%d")
        
        base_name = f"{report_name}_{region_name}_{date_suffix}"
        extension = "xlsx"
        
        download_file_path = config['download_dir'] / f"{base_name}.{extension}"
        counter = 1
        while download_file_path.exists():
            download_file_path = config['download_dir'] / f"{base_name}v{counter}.{extension}"
            counter += 1
        
        await download.save_as(str(download_file_path))
        print(f"[DEBUG] 文件下载完成: {download_file_path.name}")

    download_event_handler = handle_download
    page.on("download", download_event_handler)

    print("[DEBUG] 点击导出按钮")
    export_selectors = ['input:has-text("导出")', 'button:has-text("导出")', '[class*="export"]']
    export_found = False
    for selector in export_selectors:
        if await page.locator(selector).count() > 0:
            await page.click(selector)
            export_found = True
            break
    
    if not export_found:
        raise Exception("未找到导出按钮")

    print("[DEBUG] 选择 Excel 格式")
    try:
        await page.wait_for_selector('span:has-text("Excel")', timeout=5000)
        await page.click('span:has-text("Excel")')
    except PlaywrightTimeoutError:
        raise Exception("未找到 Excel 选项")

    print("[DEBUG] 等待在线导出按钮")
    online_export_selector = 'input[value="在线导出"]'
    
    try:
        await page.wait_for_selector(online_export_selector, timeout=10000)
        await page.click(online_export_selector)
    except PlaywrightTimeoutError:
        try:
            await page.wait_for_selector('input[atp="baseDialog_btnOK"]', timeout=5000)
            await page.click('input[atp="baseDialog_btnOK"]')
        except PlaywrightTimeoutError:
            raise Exception("未找到在线导出按钮")

    print("[INFO] 等待下载完成...")
    download_timeout = 180
    start_time = time.time()
    
    while time.time() - start_time < download_timeout:
        if download_file_path and download_file_path.exists():
            break
        await asyncio.sleep(1)
        if int(time.time() - start_time) % 30 == 0:
            print(f"[INFO] 等待中... ({int(time.time() - start_time)}秒)")
    
    if not download_file_path or not download_file_path.exists():
        raise Exception(f"下载超时！报表未成功下载")

    page.remove_listener("download", download_event_handler)
    print(f"[SUCCESS] 报表导出完成: {download_file_path.name}")
    return download_file_path


async def main():
    print("=" * 70)
    print("测试导出功能 - 阶段1")
    print("=" * 70)

    # 计算默认月份
    today = datetime.date.today()
    if today.month == 1:
        start_month = f"{today.year - 1}-01"
        end_month = f"{today.year - 1}-12"
    else:
        start_month = f"{today.year}-01"
        end_month = f"{today.year}-{str(today.month - 1).zfill(2)}"

    print("\n可用的区域配置:")
    for i, config in enumerate(REGION_CONFIGS):
        print(f"  {i+1}. {config['name']}")
        print(f"     包含区域: {', '.join(config['regions'])}")
    
    selected_region_indices = input("\n请输入要测试的区域序号（多个用逗号分隔，默认全导）: ").strip()
    
    if selected_region_indices:
        parts = selected_region_indices.replace('，', ',').split(',')
        region_indices = [int(i.strip()) - 1 for i in parts if i.strip().isdigit()]
        selected_regions = [REGION_CONFIGS[i] for i in region_indices if 0 <= i < len(REGION_CONFIGS)]
    else:
        selected_regions = REGION_CONFIGS.copy()  # 默认全导
    
    print(f"\n将测试以下区域: {', '.join([r['name'] for r in selected_regions])}")
    
    config = load_config()
    results = []
    
    playwright = await async_playwright().start()
    browser = await playwright.chromium.launch(
        headless=False,
        args=["--start-maximized", "--disable-popup-blocking"]
    )
    context = await browser.new_context(viewport={"width": 1920, "height": 1080})
    page = await context.new_page()
    
    page.on("dialog", lambda dialog: dialog.accept())

    try:
        await login(page, config)
        
        print(f"\n{'='*60}")
        print(f"访问报表: {REPORT_CONFIG['name']}")
        print(f"{'='*60}")
        
        print(f"[DEBUG] 正在访问报表页面: {REPORT_CONFIG['url']}")
        try:
            await page.goto(REPORT_CONFIG['url'], wait_until="domcontentloaded", timeout=60000)
            print("[DEBUG] 页面DOM加载完成")
            
            # 等待网络空闲
            await page.wait_for_load_state("networkidle", timeout=60000)
            print(f"[INFO] 报表页面加载完成")
        except PlaywrightTimeoutError:
            print("[WARNING] 页面加载超时，尝试继续执行")
            pass
        
        # 等待报表默认数据加载完成
        print("[INFO] 等待报表默认数据加载完成...")
        await asyncio.sleep(5)
        
        for region_config in selected_regions:
            try:
                print(f"\n{'='*50}")
                print(f"测试区域: {region_config['name']}")
                print(f"{'='*50}")
                
                # 如果是"整体"区域，跳过筛选直接导出（报表默认就是全选）
                if region_config['name'] == '整体':
                    print("[DEBUG] 整体区域，跳过筛选直接导出")
                else:
                    await select_region_filter(page, region_config['regions'])
                    print("[DEBUG] 区域筛选完成，等待2秒...")
                    await asyncio.sleep(2)
                
                # 刷新报表
                print("[DEBUG] 开始刷新报表...")
                await refresh_report(page)
                # 等待数据刷新完成
                print("[INFO] 等待报表数据刷新完成...")
                await asyncio.sleep(8)
                
                download_path = await export_report(page, config, REPORT_CONFIG['name'], region_config['name'])
                results.append({
                    'region': region_config['name'],
                    'success': True,
                    'file': download_path.name
                })
                
                await asyncio.sleep(2)
                
            except Exception as e:
                print(f"[ERROR] {region_config['name']} 导出失败: {e}")
                results.append({
                    'region': region_config['name'],
                    'success': False,
                    'error': str(e)
                })

    finally:
        await context.close()
        await browser.close()
        await playwright.stop()
    
    # 输出测试结果
    print("\n" + "="*60)
    print("测试结果汇总")
    print("="*60)
    
    success_count = sum(1 for r in results if r['success'])
    total_count = len(results)
    
    print(f"\n总测试: {total_count} | 成功: {success_count} | 失败: {total_count - success_count}")
    
    print("\n详细结果:")
    for result in results:
        if result['success']:
            print(f"  ✓ {result['region']}: {result['file']}")
        else:
            print(f"  ✗ {result['region']}: {result['error']}")
    
    print(f"\n[INFO] 下载文件存储路径: {DOWNLOAD_DIR}")


if __name__ == "__main__":
    import datetime
    asyncio.run(main())
