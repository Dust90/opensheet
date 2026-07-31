import { expect, test } from "@playwright/test";

test("OpenSheet demo loads", async ({ page }) => {
  await page.goto("/");
  await expect(page.locator("#root")).toBeVisible();
});
