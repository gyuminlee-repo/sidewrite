// Some tests run a Python tool that needs a package outside the standard
// library (tools/build_marked_diff.py needs lxml, the corpus builder needs
// python-docx). Where the package is not installed (CI installs none) such a
// test is skipped with the reason, as the Python tests of the tool are.
//   test("...", { skip: missing("lxml") }, async () => { ... })
import { execFileSync } from "node:child_process";

const known = new Map();

// false when python3 can import every module, else the reason for the skip.
export function missing(...modules) {
  const key = modules.join(",");
  if (!known.has(key)) {
    let reason = false;
    try {
      execFileSync(process.env.CC_OFFICE_PYTHON || "python3", ["-c", `import ${key}`], {
        stdio: "ignore",
      });
    } catch {
      reason = `python3 cannot import ${key}`;
    }
    known.set(key, reason);
  }
  return known.get(key);
}
