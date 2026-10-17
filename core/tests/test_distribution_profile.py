"""交付范围只控制显示；完整版和 S1/S2 版共用同一业务实现。"""

from __future__ import annotations

import json
import importlib.util
from pathlib import Path
import re
import shutil
import subprocess

import pytest


ROOT = Path(__file__).resolve().parents[2]
NODE = shutil.which("node")
FRONTEND = ROOT / "core/frontend/web/index.html"


@pytest.mark.parametrize("shell", ["shell-flask", "shell-pywebview"])
@pytest.mark.parametrize("profile_id", ["full", "s1s2"])
def test_external_reference_and_its_default_are_shipped_without_changing_source(tmp_path, monkeypatch, shell, profile_id):
    manifest = json.loads((ROOT / "版本管理/VERSION_MANIFEST.json").read_text(encoding="utf-8"))
    filename = Path(manifest["configs"]["S2-CFG-02"]["file"]).name
    private = tmp_path / "config-real"
    (private / "默认配置").mkdir(parents=True)
    (private / filename).write_bytes(b"private-current-template")
    (private / "默认配置" / filename).write_bytes(b"independent-default-template")
    (private / "3.0-test.xlsx").write_bytes(b"s3-only")
    before = {p.relative_to(private): p.read_bytes() for p in private.rglob("*.xlsx")}
    monkeypatch.syspath_prepend(str(ROOT / "shell-flask"))
    spec = importlib.util.spec_from_file_location(f"external_pack_{shell.replace('-', '_')}", ROOT / shell / "build_exe.py")
    builder = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(builder)
    monkeypatch.setattr(builder, "CONFIG_TEMPLATE", private)
    release = tmp_path / "release"
    if shell == "shell-flask":
        monkeypatch.setattr(builder, "DIST", release)
        builder.assemble_config_template()
    else:
        builder.assemble_config_template(release)
    from distribution_profile import DEFAULT_CONFIG, filter_release_config, load_config
    filter_release_config(release / "config", load_config(DEFAULT_CONFIG), profile_id)
    assert (release / "config" / filename).read_bytes() == before[Path(filename)]
    assert (release / "config/默认配置" / filename).read_bytes() == before[Path("默认配置") / filename]
    assert (release / "config/3.0-test.xlsx").exists() is (profile_id == "full")
    assert {p.relative_to(private): p.read_bytes() for p in private.rglob("*.xlsx")} == before


def test_linux_public_payload_keeps_external_headers_and_removes_private_rows(tmp_path):
    import sys
    from openpyxl import load_workbook
    script = (ROOT / "shell-flask/Linux-2-打包DEB与TARGZ.sh").read_text(encoding="utf-8")
    block = re.search(r"<<'PY'\n(import shutil\nimport sys\nimport json\n.*?\n)PY\n", script, re.S)
    assert block is not None
    manifest_path = ROOT / "版本管理/VERSION_MANIFEST.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    filename = Path(manifest["configs"]["S2-CFG-02"]["file"]).name
    source = tmp_path / "source"
    shutil.copytree(ROOT / "config", source)
    book = load_workbook(source / filename)
    headers = [c.value for c in book["集中系统数据"][2]]
    book["集中系统数据"]["B3"] = "private-bank"
    book["集中系统数据"]["C3"] = 17
    book["参照表"].append(["private-code", "private-bank", "private-region"])
    book.save(source / filename)
    book.close()
    before = (source / filename).read_bytes()
    target = tmp_path / "release/config"
    (target / "默认配置").mkdir(parents=True)
    result = subprocess.run([sys.executable, "-X", "utf8", "-", str(source), str(target), str(manifest_path)], input=block[1], text=True, encoding="utf-8", capture_output=True)
    assert result.returncode == 0, result.stderr
    for path in (target / filename, target / "默认配置" / filename):
        exported = load_workbook(path)
        assert [c.value for c in exported["集中系统数据"][2]] == headers
        assert exported["集中系统数据"].max_row == 2
        assert exported["参照表"].max_row == 1
        exported.close()
    assert (source / filename).read_bytes() == before


@pytest.mark.parametrize("profile_id", ["full", "s1s2"])
def test_windows_build_finalizes_with_same_profile_as_linux(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, profile_id: str
) -> None:
    """两个 Windows 构建均覆盖前端后，必须仍写入与 Linux 相同的交付方案。"""
    monkeypatch.syspath_prepend(str(ROOT / "shell-flask"))
    spec = importlib.util.spec_from_file_location(
        "windows_delivery_build_test", ROOT / "shell-flask/build_windows_all.py"
    )
    assert spec is not None and spec.loader is not None
    builder = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(builder)
    monkeypatch.setattr(builder, "ROOT", tmp_path)
    monkeypatch.setattr(builder, "DIST_ROOT", tmp_path / "dist")
    commands = []
    source = FRONTEND.read_bytes()

    def run_builder(command, *, cwd, env, stdin):
        commands.append(command)
        output = Path(env["FLASK_DIST_DIR"])
        win7 = Path(command[-1]).name == "build-win7.bat"
        (output / (builder.executable_stem("flask", win7=win7) + ".exe")).touch()
        # Win7 跟现代版一样会重新复制完整前端，方案必须在两者完成后才注入。
        page = output / "core/frontend/web/index.html"
        page.parent.mkdir(parents=True, exist_ok=True)
        page.write_bytes(source)
        (output / "config").mkdir(exist_ok=True)
        if win7:
            (output / "运行库修复").mkdir()
            (output / "运行环境说明-Win7.txt").touch()
        return subprocess.CompletedProcess(command, 0)

    monkeypatch.setattr(builder.subprocess, "run", run_builder)
    assert builder.main(["--profile", profile_id]) == 0
    outputs = list((tmp_path / "dist").iterdir())
    assert len(outputs) == 1
    output = outputs[0]
    assert len(commands) == 2
    assert [Path(command[-1]).name for command in commands] == [
        "build-modern.bat", "build-win7.bat"
    ]
    assert not (output / "BUILD_INCOMPLETE.txt").exists()
    assert (output / "distribution-profile.txt").read_text().strip() == profile_id
    assert output.name.endswith("_S1-S2") is (profile_id == "s1s2")

    from distribution_profile import DEFAULT_CONFIG, embed_profile, load_config

    linux_page = tmp_path / "linux-page.html"
    linux_page.write_bytes(source)
    embed_profile(linux_page, load_config(DEFAULT_CONFIG), profile_id)
    windows_page = output / "core/frontend/web/index.html"
    assert windows_page.read_bytes() == linux_page.read_bytes()
    assert FRONTEND.read_bytes() == source
    payload = json.loads(re.search(
        r'<script id="audit-distribution-config" type="application/json">(.*?)</script>',
        windows_page.read_text(encoding="utf-8"), re.S,
    ).group(1))
    assert payload["enabledSystems"] == (
        ["S1", "S2"] if profile_id == "s1s2" else ["S1", "S2", "S3", "S4"]
    )


@pytest.mark.skipif(NODE is None, reason="未安装 Node.js，无法执行前端行为检查")
def test_delivery_profiles_render_only_their_visible_systems(tmp_path: Path) -> None:
    html = FRONTEND.read_text(encoding="utf-8")
    # 实际页面脚本须仍能被浏览器解析。
    inline = "\n".join(re.findall(r"<script(?:\s[^>]*)?>(.*?)</script>", html, re.S))
    script = tmp_path / "frontend.js"
    script.write_text(inline, encoding="utf-8")
    subprocess.run([NODE, "--check", str(script)], check=True, capture_output=True, text=True)

    def section(start: str, end: str) -> str:
        return html.split(start, 1)[1].split(end, 1)[0]

    functions = "function isS1S2Edition" + section(
        "function isS1S2Edition", "function updateAuditWorkbenchMode"
    )
    functions += "function renderGeneral" + section(
        "function renderGeneral", "/* ---- 高级设置页签"
    )
    functions += "function renderGuide" + section(
        "function renderGuide", "/* ---- 常见问题列表"
    )
    flows = "const GUIDE_FLOWS=" + section("const GUIDE_FLOWS=", ";") + ";"
    faq = json.loads((ROOT / "docs/常见问题.json").read_text(encoding="utf-8"))
    program = r"""
const assert=require('node:assert/strict');
let reduced=false;
const pane={innerHTML:''};
const state={};
const document={body:{classList:{contains:()=>reduced}},getElementById:()=>pane};
const esc=value=>String(value);
const configDefaultStatus=()=>'';
const centralConfigDescription=()=>'';
const syncGeneralControls=()=>{};
const auditExtrasUnlocked=()=>true;
""" + functions + flows + "\nconst faq=" + json.dumps(faq, ensure_ascii=False) + r""";
const entries=[{date:'2026-10-15',groups:[{label:'优化',items:['S1 逐笔审核','S3 大集中统计系统','S4 地金系统','报表采集外部核对']}]}];
assert.equal(editionFaqItems(faq),faq);
assert.equal(editionUpdateEntries(entries),entries);
renderGeneral();
assert.match(pane.innerHTML,/大集中通用配置/);
renderGuide();
assert.match(pane.innerHTML,/本期数值核对/);
reduced=true;
const visibleFaq=editionFaqItems(faq);
assert.ok(visibleFaq.length>0&&visibleFaq.length<faq.length);
assert.ok(visibleFaq.some(item=>item.question==='第一次使用，我需要在界面上做哪些准备？'));
assert.ok(visibleFaq.some(item=>item.answer.includes('大集中数据文件')));
for(const item of visibleFaq){
  assert.equal(includesUndeliveredSystemInfo(item.question+item.answer),false);
  assert.doesNotMatch(item.answer,/六册|三个系统/);
}
assert.deepEqual(editionUpdateEntries(entries)[0].groups[0].items,['S1 逐笔审核','报表采集外部核对']);
assert.equal(entries[0].groups[0].items.length,4);
assert.deepEqual(editionUpdateEntries([{groups:[{items:['S3 功能']}]}]),[]);
assert.equal(includesUndeliveredSystemInfo('升级后检查 3.3「报表清单」'),true);
assert.equal(includesUndeliveredSystemInfo('整包版本 V26.1.3.1'),false);
assert.deepEqual(editionUpdateEntries([{groups:[{items:['大集中指标查询本期/上期']}]}]),[]);
renderGeneral();
assert.match(pane.innerHTML,/逐笔统计系统配置/);
assert.match(pane.innerHTML,/报表采集系统配置/);
assert.doesNotMatch(pane.innerHTML,/大集中通用配置|大集中执行比较配置|大集中本期数值核对配置|大集中指标比较拆分配置/);
renderGuide();
assert.doesNotMatch(pane.innerHTML,/本期数值核对|两期数据 ⇨ 比较文件|机构说明 ⇨ 说明文件/);
const numbers=[...pane.innerHTML.matchAll(/class="guide-flow-num">(\d+)/g)].map(match=>Number(match[1]));
assert.deepEqual(numbers,numbers.map((_,index)=>index+1));
console.log('profile behavior passed');
"""
    result = subprocess.run(
        [NODE, "-e", program], check=True, capture_output=True, text=True, encoding="utf-8"
    )
    assert "profile behavior passed" in result.stdout


@pytest.mark.skipif(NODE is None, reason="未安装 Node.js，无法执行前端行为检查")
def test_xml_tool_notes_use_one_collapse_mechanism() -> None:
    """说明使用统一卡片折叠，避免 details 原生切换与 collapsed 样式相互抵消。"""
    html = FRONTEND.read_text(encoding="utf-8")
    render = "function renderAdvanced" + html.split(
        "function renderAdvanced", 1
    )[1].split("/* 高级设置双布局", 1)[0]
    listener = html.split("/* 设置中心卡片标题条", 1)[1].split(
        "/* ---- 更新与答疑", 1
    )[0].split("*/", 1)[1]
    program = r"""
const assert=require('node:assert/strict');
let click;
const pane={innerHTML:'',addEventListener:(event,fn)=>{if(event==='click')click=fn;}};
const document={getElementById:()=>pane};
const state={};
const esc=value=>String(value);
const advancedLayoutMode=()=> 'system';
const envGroupCollapsed=()=>true,envZoneCollapsed=()=>true;
const isSystemVisible=()=>true,syncGeneralControls=()=>{};
""" + render + listener + r"""
renderAdvanced();
const notes=pane.innerHTML.match(/<div class="gen-card"><p class="gen-card-title">处理工具与兼容性说明<\/p>([\s\S]*?)<\/div><\/div>/);
assert.ok(notes,'说明应默认展开，使用普通卡片而非原生 details');
assert.match(notes[1],/zipfile/);
assert.doesNotMatch(pane.innerHTML,/<details class="gen-card"/);
const classes=new Set(['gen-card']);
const card={classList:{contains:name=>classes.has(name),toggle:name=>{
  if(classes.has(name)){classes.delete(name);}else{classes.add(name);}
}}};
const title={parentElement:card};
const event={target:{closest:()=>title}};
assert.equal(classes.has('collapsed'),false);
click(event);assert.equal(classes.has('collapsed'),true);
click(event);assert.equal(classes.has('collapsed'),false);
"""
    result = subprocess.run([NODE, "-e", program], capture_output=True, text=True, encoding="utf-8")
    assert result.returncode == 0, result.stderr


def test_distribution_unlock_password_and_process_lifetime(tmp_path: Path) -> None:
    from base_audit.web_app import WebApi

    api = WebApi(tmp_path)
    paths = [tmp_path / "data" / name for name in ("开发设置.json", "用户设置.json")]
    original = {path: path.read_bytes() for path in paths if path.exists()}
    assert api.get_state()["distributionUnlocked"] is False
    developer_access = api.get_state()["developerAccess"]
    assert all(developer_access.values())
    for password in ("", "HZPBC", "hzpbc790", None):
        assert api.unlock_distribution_ui(password)["ok"] is False
        assert api.get_state()["distributionUnlocked"] is False
        assert api.get_state()["developerAccess"] == developer_access
    result = api.unlock_distribution_ui("HZPBC790")
    assert result["ok"] is True
    assert result["state"]["distributionUnlocked"] is True
    assert all(result["state"]["developerAccess"].values())
    assert api.get_state()["distributionUnlocked"] is True
    assert {path: path.read_bytes() for path in original} == original
    restarted = WebApi(tmp_path)
    assert restarted.get_state()["distributionUnlocked"] is False
    assert restarted.get_state()["developerAccess"] == developer_access


@pytest.mark.skipif(NODE is None, reason="未安装 Node.js，无法执行前端行为检查")
def test_about_five_click_unlock_restores_embedded_profile_ui(tmp_path: Path) -> None:
    html = FRONTEND.read_text(encoding="utf-8")

    def section(start: str, end: str) -> str:
        return start + html.split(start, 1)[1].split(end, 1)[0]

    functions = section("function isS1S2Edition", "function updateAuditWorkbenchMode")
    functions += section("let distributionUnlockClicks=", "let developerUnlockTarget=")
    functions += section("let developerUnlockTarget=", "let pendingConfirm=")
    config = json.loads((ROOT / "shell-flask/distribution-profiles.json").read_text(encoding="utf-8"))
    payload = {"enabledSystems": config["profiles"]["s1s2"]["enabledSystems"], "systems": config["systems"]}
    program = r"""
const assert=require('node:assert/strict');
let state={}, now=0, settingsRenders=0, accept=false;
const Date={now:()=>now};
const currentPage='settings', CFG={data:{},tab:'guide'};
const setSettingsTab=()=>settingsRenders++;
const switchPage=()=>{throw new Error('must remain on current page')};
const nodes={};
const element=id=>nodes[id]||(nodes[id]={style:{display:'none'},value:'',textContent:'',focus:()=>{}});
const classes=new Set();
const nav=['S1','S2','S3','S4'].map(system=>({dataset:{system},hidden:false}));
const document={activeElement:{focus:()=>{}},body:{dataset:{},classList:{
  contains:value=>classes.has(value),toggle:(value,on)=>on?classes.add(value):classes.delete(value)
}},getElementById:element,querySelectorAll:()=>nav};
element('audit-distribution-config').textContent=PAYLOAD;
const bridge={api:{unlock_distribution_ui:async password=>{
  assert.equal(password,'typed-password');
  return accept?{ok:true,state:{distributionUnlocked:true,developerAccess:{S1:true,S2:true,S3:true,S4:true}}}:{ok:false,message:'wrong password'};
}}};
""".replace("PAYLOAD", json.dumps(json.dumps(payload, ensure_ascii=False), ensure_ascii=False))
    program += functions + r"""
(async()=>{
applyDistributionProfile();
assert.equal(nav[2].hidden,true);
openAbout();
for(let i=0;i<4;i++)registerDistributionUnlockClick();
assert.equal(element('developerUnlockMask').style.display,'none');
now=2100;
registerDistributionUnlockClick();
assert.equal(element('developerUnlockMask').style.display,'none');
for(let i=0;i<4;i++)registerDistributionUnlockClick();
assert.equal(element('developerUnlockMask').style.display,'flex');
assert.equal(element('developerPassword').value,'');
element('developerPassword').value='typed-password';
const pending=forceDeveloperEntry();
closeDeveloperUnlock();
assert.equal(element('developerUnlockMask').style.display,'flex');
await pending;
assert.equal(element('developerUnlockError').textContent,'wrong password');
assert.equal(nav[2].hidden,true);
closeDeveloperUnlock();
assert.equal(element('developerPassword').value,'');
assert.equal(isSystemVisible('S3'),false);
openAbout();
for(let i=0;i<5;i++)registerDistributionUnlockClick();
element('developerPassword').value='typed-password';
accept=true;
await forceDeveloperEntry();
assert.equal(element('developerUnlockMask').style.display,'none');
assert.ok(nav.every(item=>!item.hidden));
assert.ok(classes.has('distribution-ui-unlocked'));
assert.equal(settingsRenders,1);
assert.equal(isSystemVisible('S3'),true);
const faq=[{category:'大集中统计系统',question:'S3'}];
assert.equal(editionFaqItems(faq),faq);
assert.deepEqual(DISTRIBUTION_CONFIG.enabledSystems,['S1','S2']);
state={};
DISTRIBUTION_CONFIG.enabledSystems=['S1','S2','S3','S4'];
openAbout();
for(let i=0;i<5;i++)registerDistributionUnlockClick();
assert.equal(element('developerUnlockMask').style.display,'none');
console.log('about unlock passed');
})().catch(error=>{console.error(error);process.exit(1)});
"""
    script = tmp_path / "about_unlock_test.js"
    script.write_text(program, encoding="utf-8")
    result = subprocess.run([NODE, str(script)], capture_output=True, text=True, encoding="utf-8")
    assert result.returncode == 0, result.stderr
    assert "about unlock passed" in result.stdout
