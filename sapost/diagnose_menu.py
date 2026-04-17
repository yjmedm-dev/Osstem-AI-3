"""
FBL5N 결과 화면 메뉴바 구조 확인
FBL5N 조회 결과 화면이 열린 상태에서 실행하세요.

실행:
  python sapost/diagnose_menu.py
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import win32com.client


def explore_menu(menu, depth=0):
    indent = "  " * depth
    try:
        menu_id   = menu.Id
        menu_text = menu.Text
        print(f"{indent}[{menu_id}]  text={menu_text}")
    except Exception as e:
        print(f"{indent}(읽기 실패: {e})")
        return
    try:
        count = menu.Children.Count
        for i in range(count):
            explore_menu(menu.Children.ElementAt(i), depth + 1)
    except Exception:
        pass


def print_wnd_children(ctrl, depth=0):
    indent = "  " * depth
    try:
        ctrl_id   = ctrl.Id
        ctrl_type = ctrl.Type
        ctrl_text = ""
        try:
            ctrl_text = ctrl.Text[:60] if ctrl.Text else ""
        except Exception:
            pass
        print(f"{indent}[{ctrl_type}] {ctrl_id}  text={ctrl_text}")
    except Exception as e:
        print(f"{indent}(읽기 실패: {e})")
        return
    try:
        count = ctrl.Children.Count
        for i in range(count):
            print_wnd_children(ctrl.Children.ElementAt(i), depth + 1)
    except Exception:
        pass


def main():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application  = sap_gui_auto.GetScriptingEngine
    connection   = application.Children(0)
    session      = connection.Children(0)

    print(f"현재 트랜잭션: {session.Info.Transaction}")

    # ── 메뉴바 전체 구조 출력 ──
    print("\n=== 메뉴바 구조 ===")
    try:
        mbar = session.findById("wnd[0]/mbar")
        explore_menu(mbar)
    except Exception as e:
        print(f"메뉴바 읽기 실패: {e}")

    # ── btn[44] 스프레드시트 버튼 시도 ──
    print("\n=== btn[44] 클릭 시도 (스프레드시트) ===")
    try:
        session.findById("wnd[0]/tbar[1]/btn[44]").press()
        print("btn[44] 클릭 성공")
        time.sleep(2)
    except Exception as e:
        print(f"btn[44] 실패: {e}")

    # ── 이후 열린 창 확인 ──
    print("\n=== 클릭 후 열린 창 (wnd[1] ~ wnd[3]) ===")
    for wnd_idx in [1, 2, 3]:
        try:
            wnd = session.findById(f"wnd[{wnd_idx}]")
            print(f"\nwnd[{wnd_idx}] title={wnd.Text}")
            print_wnd_children(wnd)
        except Exception as e:
            print(f"wnd[{wnd_idx}] 없음")


if __name__ == "__main__":
    main()
