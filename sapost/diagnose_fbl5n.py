"""
FBL5N 선택 화면의 모든 컨트롤 ID를 출력하는 진단 스크립트
SAP GUI가 FBL5N 선택 화면에 열려 있는 상태에서 실행하세요.

실행:
  python sapost/diagnose_fbl5n.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import win32com.client


def print_children(ctrl, depth=0):
    indent = "  " * depth
    try:
        ctrl_id   = ctrl.Id
        ctrl_type = ctrl.Type
        ctrl_name = getattr(ctrl, "Name", "")
        ctrl_text = ""
        try:
            ctrl_text = ctrl.Text[:40] if ctrl.Text else ""
        except Exception:
            pass
        print(f"{indent}[{ctrl_type}] {ctrl_id}  name={ctrl_name}  text={ctrl_text}")
    except Exception as e:
        print(f"{indent}(읽기 실패: {e})")
        return

    try:
        count = ctrl.Children.Count
        for i in range(count):
            print_children(ctrl.Children.ElementAt(i), depth + 1)
    except Exception:
        pass


def main():
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application  = sap_gui_auto.GetScriptingEngine
    connection   = application.Children(0)
    session      = connection.Children(0)

    print(f"\n현재 트랜잭션: {session.Info.Transaction}")
    print("=" * 70)
    print("FBL5N 선택 화면 컨트롤 목록:")
    print("=" * 70)

    try:
        usr = session.findById("wnd[0]/usr")
        print_children(usr)
    except Exception as e:
        print(f"오류: {e}")
        print("FBL5N 선택 화면이 열려 있는지 확인하세요.")


if __name__ == "__main__":
    main()
