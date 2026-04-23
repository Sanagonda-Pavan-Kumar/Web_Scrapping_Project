const { chromium } = require("playwright");
const fs = require("fs");
const xlsx = require("xlsx"); // Import the xlsx package for Excel export

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
    await page.selectOption("#pagination", "500");
    await page.waitForTimeout(2000); // Allow time for the page to load
  }

  async function extractMemberDetails() {
    return await page.evaluate(() => {
      const getText = (header) => {
        const headingElement = Array.from(document.querySelectorAll("h4")).find(
          (h) => h.innerText.trim() === header
        );
        return headingElement
          ? headingElement.closest(".location-heading-brand").querySelector("p").innerText.trim()
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
  }

  try {
    await navigateToMainPage();

    // Get all "View Details" links
    const detailLinks = await page.$$eval(
      ".view-details-btn a",
      (links) => links.map((link) => link.href)
    );

    const allMemberData = [];

    for (const link of detailLinks) {
      // Navigate to each member's details page
      await page.goto(link, { waitUntil: "domcontentloaded" });

      // Extract the company name
      const companyName = await page.evaluate(() => {
        const companyNameElement = document.querySelector("div.container-xxl h2");
        return companyNameElement ? companyNameElement.innerText.trim() : "Company name not found";
      });

      // Extract the member data
      const memberData = await extractMemberDetails();
      console.log("Extracted Member Data:", { CompanyName: companyName, ...memberData });

      // Save the member data
      allMemberData.push({
        CompanyName: companyName,
        ...memberData,
      });

      // Go back to the main directory page
      await navigateToMainPage();
    }

    // Write all data to a JSON file
    fs.writeFileSync("allMembersData.json", JSON.stringify(allMemberData, null, 2));

    // Write all data to an Excel file
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.json_to_sheet(allMemberData);
    xlsx.utils.book_append_sheet(workbook, worksheet, "Members");
    xlsx.writeFile(workbook, "allMembersData.xlsx");

    console.log("Data extracted and saved to allMembersData.json and allMembersData.xlsx");

  } catch (error) {
    console.error("Error:", error.message);
  } finally {
    await browser.close();
  }
})();
