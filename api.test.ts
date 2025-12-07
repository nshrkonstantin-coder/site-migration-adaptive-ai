import { expect } from "expect";
import { health } from "./api";

async function testHealth() {
  const res = await health();
  expect(res.ok).toBe(true);
  expect(typeof res.name).toBe("string");
  expect(res.name.length).toBeGreaterThan(0);
}

type TestResult = {
  passedTests: string[];
  failedTests: { name: string; error: string }[];
};

export async function _runApiTests() {
  const result: TestResult = { passedTests: [], failedTests: [] };
  const tests = [testHealth];

  for (const t of tests) {
    try {
      await t();
      result.passedTests.push(t.name);
    } catch (e) {
      result.failedTests.push({
        name: t.name,
        error: e instanceof Error ? e.message : "Unknown error",
      });
    }
  }

  return result;
}
