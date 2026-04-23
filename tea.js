const { chromium } = require("playwright");
const fs = require("fs");
const xlsx = require("xlsx");

(async () => {
  const browser = await chromium.launch({
    headless: false,
    args: ["--start-maximized"],
  });

  const context = await browser.newContext({
    viewport: null,
    userAgent:
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
    extraHTTPHeaders: {
      "Accept-Language": "en-US,en;q=0.9",
      Referer: "https://ieema.org/",
    },
  });

  const page = await context.newPage();

  async function navigateToMainPage() {
    await page.goto("https://ieema.org/member-directory/", {
      waitUntil: "domcontentloaded",
      timeout: 60000,
    });

    await page.waitForSelector("#pagination", { timeout: 40000 });

    // Change pagination to 500 and wait for reload properly
    await Promise.all([
      page.waitForLoadState("networkidle"),
      page.selectOption("#pagination", "500"),
    ]);

    await page.waitForSelector(".view-details-btn a");
  }

  try {
    await navigateToMainPage();

    // Get all detail links
    const detailLinks = await page.$$eval(
      ".view-details-btn a",
      (links) => links.map((link) => link.href)
    );

    console.log(`Total links found: ${detailLinks.length}`);

    const allMemberData = [];

    // Loop through each link (open in new tab)
    for (const link of detailLinks) {
      const detailPage = await context.newPage();

      try {
        await detailPage.goto(link, {
          waitUntil: "domcontentloaded",
          timeout: 60000,
        });

        await detailPage.waitForSelector("div.container-xxl h2");

        // Extract company name
        const companyName = await detailPage.evaluate(() => {
          const el = document.querySelector("div.container-xxl h2");
          return el ? el.innerText.trim() : "Company name not found";
        });

        // Extract member details
        const memberData = await detailPage.evaluate(() => {
          const getText = (header) => {
            const headingElement = Array.from(
              document.querySelectorAll("h4")
            ).find((h) => h.innerText.trim() === header);

            return headingElement
              ? headingElement
                  .closest(".location-heading-brand")
                  .querySelector("p").innerText.trim()
              : "";
          };

          return {
            Name: getText("Name"),
            Designation: getText("Designation"),
            Email: getText("Email").replace(/<.*?>/g, "").trim(),
            State: getText("State"),
            Region: getText("Region"),
            City: getText("City"),
          };
        });

        const finalData = {
          CompanyName: companyName,
          ...memberData,
        };

        console.log("Extracted:", finalData);

        allMemberData.push(finalData);
      } catch (err) {
        console.error(`Failed for ${link}:`, err.message);
      } finally {
        await detailPage.close(); // always close tab
      }
    }

    // Save JSON
    fs.writeFileSync(
      "allMembersData.json",
      JSON.stringify(allMemberData, null, 2)
    );

    // Save Excel
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.json_to_sheet(allMemberData);
    xlsx.utils.book_append_sheet(workbook, worksheet, "Members");
    xlsx.writeFile(workbook, "allMembersData.xlsx");

    console.log("✅ Data saved to JSON & Excel");
  } catch (error) {
    console.error("❌ Error:", error.message);
  } finally {
    await browser.close();
  }
})();