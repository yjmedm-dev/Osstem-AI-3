"""
ZQSAB01 선택 화면의 모든 컨트롤 ID를 출력하는 진단 스크립트

사용법:
  1. SAP GUI에서 ZQSAB01 선택 화면을 열어 둡니다.
  2. 아래 명령을 실행합니다.
       python sapost/diagnose_zqsab01.py
  3. 출력된 필드 ID를 config/config.ini [ZQSAB01] 섹션에 반영합니다.

     [ZQSAB01]
     gjahr_field       = wnd[0]/usr/...  ← 회계연도 필드 ID
     from_period_field = wnd[0]/usr/...  ← 시작기간 필드 ID
     to_period_field   = wnd[0]/usr/...  ← 종료기간 필드 ID
     bukrs_field       = wnd[0]/usr/...  ← 회사코드 필드 ID
     grid_id           = wnd[0]/usr/...  ← 결과 ALV 그리드 ID
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import win32com.client


def print_children(ctrl, depth: int = 0):
    indent = "  " * depth
    try:
        cid   = ctrl.Id
        ctype = ctrl.Type
        cname = getattr(ctrl, "Name", "")
        ctext = ""
        try:
            raw = ctrl.Text
            ctext = (raw[:50] + "…") if raw and len(raw) > 50 else (raw or "")
        except Exception:
            pass
        print(f"{indent}[{ctype}]  {cid}  name={cname}  text={ctext}")
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

    current_tx = session.Info.Transaction
    print(f"\n현재 트랜잭션: {current_tx}")

    if current_tx.upper() != "ZQSAB01":
        print(
            "\n[경고] ZQSAB01 선택 화면이 열려 있지 않습니다.\n"
            "SAP GUI에서 ZQSAB01을 실행한 후 다시 이 스크립트를 실행하세요.\n"
        )

    print("=" * 70)
    print("ZQSAB01 선택 화면 컨트롤 목록")
    print("=" * 70)

    try:
        usr = session.findById("wnd[0]/usr")
        print_children(usr)
    except Exception as e:
        print(f"오류: {e}")
        print("ZQSAB01 선택 화면이 열려 있는지 확인하세요.")

    print("\n" + "=" * 70)
    print("전체 창(wnd[0]) 구조 (ALV 그리드 포함)")
    print("=" * 70)
    try:
        wnd = session.findById("wnd[0]")
        print_children(wnd)
    except Exception as e:
        print(f"오류: {e}")


if __name__ == "__main__":
    main()
