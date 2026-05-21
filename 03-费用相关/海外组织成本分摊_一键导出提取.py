import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
EXPORT_SCRIPT = BASE_DIR / "海外组织成本分摊_阶段1导出.py"
EXTRACT_SCRIPT = BASE_DIR / "海外组织成本分摊_阶段2提取.py"


def run_export() -> None:
    if not EXPORT_SCRIPT.exists():
        raise FileNotFoundError(f"未找到导出脚本: {EXPORT_SCRIPT}")

    print("=" * 70)
    print("步骤1/2：执行报表导出")
    print("=" * 70)
    print(f"[INFO] 将运行: {EXPORT_SCRIPT.name}")

    # 导出脚本本身有交互，直接透传终端输入输出
    subprocess.run([sys.executable, str(EXPORT_SCRIPT)], check=True, cwd=str(BASE_DIR))


def run_extract(month_str: str | None) -> None:
    if not EXTRACT_SCRIPT.exists():
        raise FileNotFoundError(f"未找到提取脚本: {EXTRACT_SCRIPT}")

    print("\n" + "=" * 70)
    print("步骤2/2：执行例子数提取")
    print("=" * 70)
    print(f"[INFO] 将运行: {EXTRACT_SCRIPT.name}")

    # 提取脚本首行会提示输入月份；这里自动喂入（空字符串=默认上月）
    input_text = (month_str or "") + "\n"
    subprocess.run(
        [sys.executable, str(EXTRACT_SCRIPT)],
        check=True,
        cwd=str(BASE_DIR),
        input=input_text,
        text=True,
    )


def main() -> None:
    print("=" * 70)
    print("海外组织成本分摊：一键导出+提取")
    print("=" * 70)

    month_input = input("请输入提取月份(YYYY-MM，直接回车=默认上月): ").strip()

    try:
        run_export()
        run_extract(month_input)
        print("\n[SUCCESS] 一键流程执行完成。")
    except subprocess.CalledProcessError as e:
        print(f"\n[ERROR] 子脚本执行失败，退出码: {e.returncode}")
        raise
    except Exception as e:
        print(f"\n[ERROR] 执行失败: {e}")
        raise


if __name__ == "__main__":
    main()
