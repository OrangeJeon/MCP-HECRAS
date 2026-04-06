from mcp.server.fastmcp import FastMCP
import os, re
import csv
import shutil
import win32com.client
import time 


PROJECT_PATH = r"C:\Hriver2\beforeLSB\hwangriver.prj"
OUTPUT_PATH =  r"C:\Hriver2\beforeLSB\flowrates.csv"

mcp = FastMCP("hecras_mcp")
import win32com.client
import pandas as pd

rc = win32com.client.Dispatch("RAS67.HECRASController")

#파일 열기
@mcp.tool() 
def open_project(project_path: str) -> dict:
    try:
        rc.Project_Open(project_path)
        return {"success": True, "project_path": project_path}
    except Exception as e:
        return {"success": False, "error": str(e)}

#plan 계산
@mcp.tool() 
def run_current_plan() -> dict:
    try:
        rc.Compute_CurrentPlan(None, None)
        return {"success": True, "message": "계산 완료"}
    except Exception as e:
        return {"success": False, "error": str(e)}

#연결여부 확인
@mcp.tool() 
def check_connection()->dict:
    try:
        version = rc.HECRASVersion()
        project = rc.CurrentProjectFile()
        plan = rc.CurrentPlanFile()

        return{
            "success": True,
            "hecras_version": version, 
            "com_object": str(rc),
            "current_project": project if project else "Null",
            "current_plan": plan if plan else "Null"
        }
    except Exception as e:
        return{
            "success": False,
            "hecras_version": version, 
            "com_object": str(rc),
            "current_project": project if project else "Null",
            "current_plan": plan if plan else "Null"
        }

#플랜에 연결된 flow 파일 경로 찾기
def get_flow_file_path()->str:
    plan_path = rc.CurrentPlanFile()
    if not plan_path:
        raise FileNotFoundError("플랜 없음")
    
    prj_dir = os.path.dirname(plan_path)

    prj_name = os.path.splitext(os.path.basename(plan_path))[0]

    with open(plan_path, "r", encoding='utf-8', errors ="ignore") as f:
        for line in f:
            if line.startswith("Flow File="):
                flow_filename = line.split("=", 1)[1].strip()
                if "." not in flow_filename:
                    flow_filename = f"{prj_name}.{flow_filename}"
                return os.path.join(prj_dir, flow_filename)
    raise FileNotFoundError("플랜파일은 있는데 Flow File 없음")

#현재 유량 조건 읽기
@mcp.tool() 
def get_flow_data() -> dict:
    try:
        flow_path = get_flow_file_path()

        with open(flow_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()

        profile_match = re.search(r"Number of Profiles\s*=\s*(\d+)", content)
        n_profiles = int(profile_match.group(1)) if profile_match else 1

        locations = re.findall(
    r"River Rch & RM=\s*(.+?),\s*(.+?),\s*([\d.]+)\s*\n([ \t\d.]+)\n", content, re.DOTALL
        )

        flow_data = []
        for river, reach, rs, flows_str in locations:
            flows = [float(v) for v in flows_str.split()]
            flow_data.append({
                "river": river.strip(),
                "reach": reach.strip(),
                "RS": rs.strip(),
                "flows": flows
            })
        return{
            "success" : True,
            "flow_file": flow_path,
            "n_profiles": n_profiles,
            "flow_data": flow_data
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

#steady 데이터 조회
@mcp.tool()
def get_steady_flow_data()->dict:
    try:
        flow_path = get_flow_file_path()

        with open(flow_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        profile_match = re.search(r"Number of Profiles\s*=\s*(\d+)", content)
        n_profile = int(profile_match.group(1)) if profile_match else 1

        name_match = re.search(r"Profile Names=(.+)", content)
        if name_match:
            profile_names = [n.strip()for n in name_match.group(1).split(",")]
        else:
            profile_names = [f"PF{i}" for i in range(1, n_profile+1)]

        locations = re.findall(
            r"River Rch & RM=\s*(.+?),\s*(.+?),\s*([\d.]+)\s*\n([\d\s]+)", content)
        
        flow_data = []
        for river, reach, rs, flow_str in locations:
            flows = [float(v) for v in flow_str.split()]
            profiles={
                profile_names[i]: flows[i]
                for i in range(min(len(profile_names), len(flows)))

            }
            flow_data.append({
                "river": river.strip(),
                "reach": reach.strip(),
                "RS": rs.strip(),
                "profiles": profiles
            })
        
        return{
            "success": True,
            "flow_file": flow_path,
            "n_profiles": n_profile,
            "profile_names": profile_names,
            "flow_data": flow_data
        }
    except Exception as e:
        return {"success": False, "error":str(e)}

#분석하는 코드
@mcp.tool()
def run_steady_flow_analysis(output_dir: str = None):
    flow_result = get_steady_flow_data()

    if not flow_result["success"]:
        print("flow data 없음")
        return False
    print("데이터 조회 완료")

    project_path = rc.CurrentProjectFile()
    if project_path:
        rc.Project_Open(project_path)
        time.sleep(2)

    plan_result = run_current_plan()
    if not plan_result['success']:
        print("plan 실패")
        return False
    print("plan 실행 완료")

    flow_file = flow_result["flow_file"]
    if output_dir is None:
        output_dir = os.path.dirname(flow_file)

    output_path = os.path.join(output_dir, "steady_flow_result.csv")
    profile_names = flow_result["profile_names"]
    flow_data = flow_result["flow_data"]

    with open(output_path, 'w', newline="", encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(["River", "Reach", "RS"] + profile_names)
        for row in flow_data:
            values = [row["profiles"].get(p, "") for p in profile_names]
            writer.writerow([row["river"], row["reach"], row["RS"]]+ values)
    return True

#steady 데이터 추가
@mcp.tool()
def add_steady_flow_profile(
    project_path: str,
    base_profile: str,
    multiplier: float,
    new_profile_name: str
):
    rc.Project_Open(project_path)

    n_profiles = 0
    profile_names = []
    flow_result = get_steady_flow_data()
    if not flow_result["success"]:
        raise ValueError(f"Flow 데이터 읽기 실패: {flow_result['error']}")
    
    profile_names = flow_result["profile_names"]
    n_profiles = flow_result["n_profiles"]
    
    print(f"[정보] 현재 프로파일 목록: {profile_names}")
    
    if base_profile not in profile_names:
        raise ValueError(f"기준 프로파일 '{base_profile}'을 찾을 수 없습니다. 현재 목록: {profile_names}")

    # 이미 추가된 경우 중단
    if new_profile_name in profile_names:
        raise ValueError(f"'{new_profile_name}'이 이미 존재합니다. 프로파일 목록: {profile_names}")

    base_idx = profile_names.index(base_profile)  # 0-based

    proj_dir  = os.path.dirname(project_path)
    proj_stem = os.path.splitext(os.path.basename(project_path))[0]

    flow_filename_raw = None
    with open(project_path, "r") as f:
        for line in f:
            if line.startswith("Flow File="):
                flow_filename_raw = line.strip().split("=", 1)[1].strip()
                break

    if not flow_filename_raw:
        raise ValueError("프로젝트 파일에 'Flow File=' 항목이 없습니다.")

    composed = flow_filename_raw if "." in flow_filename_raw else f"{proj_stem}.{flow_filename_raw}"
    flow_file_path = os.path.join(proj_dir, composed)

    if not os.path.exists(flow_file_path):
        raise FileNotFoundError(f"Flow 파일을 찾을 수 없습니다: {flow_file_path}")

    print(f"[정보] Flow 파일 경로: {flow_file_path}")

  
    with open(flow_file_path, "r") as f:
        lines = f.readlines()

    new_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]

        # ── Number of Profiles 수정
        if re.match(r"Number of Profiles\s*=", line):
            new_lines.append(f"Number of Profiles= {n_profiles + 1}\n")
            i += 1

        # ── Profile Names: 중복 없이 추가
        if re.match(r"Profile Names\s*=", line):
            existing_names = line.strip().split("=", 1)[1].strip()
            # 혹시 이전 실행으로 중복된 PF10 제거 후 다시 추가
            clean_names = ",".join(
                n for n in existing_names.split(",") if n != new_profile_name
            )
            new_lines.append(f"Profile Names={clean_names},{new_profile_name}\n")
            i += 1

        # ── River Rch & RM= 다음 줄이 유량값 행
        if line.startswith("River Rch & RM="):
            new_lines.append(line)
            i += 1

            all_vals = []
            val_lines = []
            while i < len(lines) and re.match(r'^[\s\d.]+$', lines[i]) and lines[i].strip():
                val_lines.append(lines[i])
                all_vals.extend(lines[i].split())
                i += 1

            try: 
                base_val = float(all_vals[base_idx])
                new_val = int(round(base_val * multiplier))
            except(IndexError, ValueError):
                new_val = 0
            
            for j, vl in enumerate(val_lines):
                if j==len(val_lines)-1:
                    new_lines.append(vl.rstrip("\n") + f"{new_val:8d}\n")
                else:
                    new_lines.append(vl)
        else:
            new_lines.append(line)
            i += 1

    # ── 4. 백업 후 저장 ───────────────────────────────────────────────────────
    backup_path = flow_file_path + ".bak"
    shutil.copy2(flow_file_path, backup_path)
    print(f"[정보] 백업 완료: {backup_path}")

    with open(flow_file_path, "w") as f:
        f.writelines(new_lines)
    rc.Project_Open(project_path)
    print(f"[완료] '{new_profile_name}' 프로파일이 추가되었습니다.")
    print(f"       기준: {base_profile}  ×  {multiplier}  ({multiplier*100:.0f}%)")
  

#명령어 파라미터 쪼개기(3개)
def parse_command(text: str, profile_names: list[str]) -> dict:
    base_profile = None
    for name in sorted(profile_names, key=len, reverse=True):
        if name in text:
            base_profile = name
            break

    multiplier = None
    m = re.search(r"(\d+(?:\.\d+)?)\s*%", text)
    if m:
        multiplier = float(m.group(1)) / 100
    else:
        m = re.search(r"(\d+(?:\.\d+)?)\s*[배xX]", text)
        if m:
            multiplier = float(m.group(1))

    new_profile_name = None
    m = re.search(r"([A-Za-z_]\w*)\s*(?:으로|로|이름으로|name)", text)
    if m and m.group(1) != base_profile:
        new_profile_name = m.group(1)
    if not new_profile_name:
        for tok in re.findall(r"[A-Za-z_]\w*", text):
            if tok not in (profile_names or []) and not re.fullmatch(r"\d+", tok):
                new_profile_name = tok
                break

    return {"base_profile": base_profile,
            "multiplier": multiplier,
            "new_profile_name": new_profile_name}

#자연어 인터페이스
@mcp.tool()
def request_add_profile(user_input: str) -> dict:
    try:
        rc.Project_Open(PROJECT_PATH)
        _, names = rc.Output_GetProfiles(0, [])
        profile_names = list(names)
        rc.QuitRas()

        parsed = parse_command(user_input, profile_names)

        missing = [k for k, v in parsed.items() if v is None]
        if missing:
            return {"success": False,
                    "message": f"다음 정보를 인식하지 못했습니다: {missing}",
                    "parsed": parsed,
                    "hint": "예) '20yr의 120%를 계산해서 PF10으로 넣어줘'"}

        return add_steady_flow_profile(
            project_path     = PROJECT_PATH,
            base_profile     = parsed["base_profile"],
            multiplier       = parsed["multiplier"],
            new_profile_name = parsed["new_profile_name"],
        )
    except Exception as e:
        return {"success": False, "error": str(e)}
    
#Profile Output Table 조회
@mcp.tool()
def open_profile_output_table(project_path: str):
    rc.ShowRAS()
    time.sleep(3)

    print(f"프로젝트 열기: {project_path}")
    rc.Project_Open(project_path)
    time.sleep(3)

    # ── TablePF: Profile Output Table 열기 ──────────────────────────
    try:
        rc.TablePF()
        print("Profile Output Table 창 열기 성공")
    except Exception as e:
        print(f"✘ TablePF 실패: {e}")

    # ── TableXS: Cross Section Output Table (필요 시 사용) ───────────
    # hec.TableXS()

    return rc

if __name__ == "__main__":
    mcp.run(transport='stdio')