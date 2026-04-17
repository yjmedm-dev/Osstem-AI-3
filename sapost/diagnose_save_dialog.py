"""
FBL5N 결과 화면에서 로컬 파일 저장 대화상자 컨트롤 ID 확인
FBL5N 조회 결과 화면이 열린 상태에서 실행하세요.

실행:
  python sapost/diagnose_save_dialog.py
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import win32com.client


def print_children(ctrl, depth=0):
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
        for i in range(ctrl.Children.Count):
            print_children(ctrl.Children.ElementAt(i), depth + 1)
    except Exception:
        pass


def main():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application  = sap_gui_auto.GetScriptingEngine
    connection   = application.Children(0)
    session      = connection.Children(0)

    print(f"현재 트랜잭션: {session.Info.Transaction}")

    # 리스트 → 스프레드시트 → 로컬 파일... 클릭
    print("\n=== 메뉴 클릭: 리스트 > 스프레드시트 > 로컬 파일... ===")
    try:
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
        print("메뉴 클릭 성공")
    except Exception as e:
        print(f"메뉴 클릭 실패: {e}")
        return

    time.sleep(2)

    # 열린 창 전체 출력
    print("\n=== 열린 대화상자 ===")
    for wnd_idx in [1, 2, 3]:
        try:
            wnd = session.findById(f"wnd[{wnd_idx}]")
            print(f"\nwnd[{wnd_idx}] title={wnd.Text}")
            print("-" * 50)
            print_children(wnd)
        except Exception:
            print(f"wnd[{wnd_idx}] 없음")


if __name__ == "__main__":
    main()
