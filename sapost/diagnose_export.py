"""
FBL5N 결과 화면에서 내보내기 버튼 클릭 후 대화상자 컨트롤 ID 확인
FBL5N 조회 결과 화면이 열린 상태에서 실행하세요.

실행:
  python sapost/diagnose_export.py
"""
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import win32com.client


def print_window(session, wnd_id):
    try:
        wnd = session.findById(wnd_id)
        print(f"\n[{wnd_id}]  title={wnd.Text}")
        print("-" * 60)
        print_children(wnd)
    except Exception as e:
        print(f"{wnd_id} 없음: {e}")


def print_children(ctrl, depth=0):
    indent = "  " * depth
    try:
        ctrl_id   = ctrl.Id
        ctrl_type = ctrl.Type
        ctrl_text = ""
        try:
            ctrl_text = ctrl.Text[:50] if ctrl.Text else ""
        except Exception:
            pass
        print(f"{indent}[{ctrl_type}] {ctrl_id}  text={ctrl_text}")
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

    print(f"현재 트랜잭션: {session.Info.Transaction}")

    # 툴바 버튼 목록 출력 (내보내기 버튼 찾기)
    print("\n=== 툴바 버튼 목록 (wnd[0]/tbar[1]) ===")
    try:
        tbar = session.findById("wnd[0]/tbar[1]")
        for i in range(tbar.Children.Count):
            btn = tbar.Children.ElementAt(i)
            try:
                print(f"  btn[{i}] id={btn.Id}  text={btn.Text}  tooltip={btn.Tooltip}")
            except Exception:
                print(f"  btn[{i}] (읽기 실패)")
    except Exception as e:
        print(f"  툴바 없음: {e}")

    # 내보내기 버튼 클릭 시도
    print("\n=== 내보내기 버튼 클릭 시도 ===")
    exported = False
    for btn_idx in [45, 46, 43, 44]:
        try:
            session.findById(f"wnd[0]/tbar[1]/btn[{btn_idx}]").press()
            print(f"  btn[{btn_idx}] 클릭 성공")
            exported = True
            time.sleep(1.5)
            break
        except Exception:
            pass

    if not exported:
        print("  툴바 버튼으로 내보내기 실패 — 메뉴 시도")
        try:
            # 메뉴: 목록 → 내보내기 → 로컬 파일
            session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
            print("  메뉴 클릭 성공")
            time.sleep(1.5)
        except Exception as e:
            print(f"  메뉴도 실패: {e}")

    # 대화상자 컨트롤 출력
    print("\n=== 대화상자 컨트롤 (wnd[1] ~ wnd[3]) ===")
    for wnd_idx in [1, 2, 3]:
        print_window(session, f"wnd[{wnd_idx}]")


if __name__ == "__main__":
    main()
