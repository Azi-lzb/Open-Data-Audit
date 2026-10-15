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
