import { describe, expect, test } from "bun:test";
import {
  daemonTaskSpec,
  hiddenLauncherScript,
  hiddenLauncherTaskSpec,
  outlookTaskSpec,
  taskXml,
} from "../src/platform/scheduler.ts";

const user = "EXAMPLE\\user";

describe("Task Scheduler XML", () => {
  test("runs at logon of the current user, interactive, least privilege, no limit", () => {
    const xml = taskXml(daemonTaskSpec("C:\\Users\\user\\AppData\\Local\\msoutlook-mcp\\bin\\msoutlook-mcp.exe", user, "C:\\Users\\user\\AppData\\Local\\msoutlook-mcp", 8787));
    expect(xml).toContain("<LogonTrigger>");
    expect(xml).toContain(`<UserId>${user}</UserId>`);
    expect(xml).toContain("<LogonType>InteractiveToken</LogonType>");
    expect(xml).toContain("<RunLevel>LeastPrivilege</RunLevel>");
    expect(xml).toContain("<ExecutionTimeLimit>PT0S</ExecutionTimeLimit>");
    expect(xml).toContain("<MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>");
    expect(xml).toContain("<RestartOnFailure>");
    expect(xml).toContain("<Command>C:\\Users\\user\\AppData\\Local\\msoutlook-mcp\\bin\\msoutlook-mcp.exe</Command>");
    expect(xml).toContain("<Arguments>serve --port 8787</Arguments>");
  });

  test("the Outlook task starts Outlook minimized through cmd", () => {
    const spec = outlookTaskSpec("C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE", user);
    const xml = taskXml(spec);
    expect(spec.exe).toBe("cmd.exe");
    expect(spec.args).toContain('/min "C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"');
    expect(xml).toContain("<Arguments>/c start &quot;&quot; /min &quot;C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE&quot;</Arguments>");
    expect(xml).toContain("<Delay>PT5S</Delay>");
  });

  test("XML special characters are escaped", () => {
    const xml = taskXml({
      description: "A & B <test>",
      exe: "C:\\a&b\\x.exe",
      args: `--name "a&b"`,
      workDir: "C:\\a<b>",
      user,
    });
    expect(xml).toContain("A &amp; B &lt;test&gt;");
    expect(xml).toContain("C:\\a&amp;b\\x.exe");
    expect(xml).toContain("&quot;a&amp;b&quot;");
    expect(xml).toContain("C:\\a&lt;b&gt;");
  });

  test("the daemon has a logon delay so Outlook starts first", () => {
    const xml = taskXml(daemonTaskSpec("x.exe", user, "C:\\"));
    expect(xml).toContain("<Delay>PT15S</Delay>");
  });

  test("the installed task uses wscript and a hidden VBS launcher", () => {
    const script = hiddenLauncherScript("C:\\Users\\user\\AppData\\Local\\msoutlook-mcp\\bin\\msoutlook-mcp.exe");
    expect(script).toContain('CreateObject("WScript.Shell")');
    expect(script).toContain("Chr(34)");
    expect(script).toContain('"C:\\Users\\user\\AppData\\Local\\msoutlook-mcp\\bin\\msoutlook-mcp.exe"');
    expect(script).toContain('& " serve"');
    expect(script).toContain("0, True");
    const xml = taskXml(hiddenLauncherTaskSpec("C:\\launcher.vbs", user, "C:\\install"));
    expect(xml).toContain("<Command>wscript.exe</Command>");
    expect(xml).toContain('<Arguments>//B //NoLogo &quot;C:\\launcher.vbs&quot;</Arguments>');
  });
});
