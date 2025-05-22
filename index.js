import cors from "cors";
import dotenv from "dotenv";
import express from "express";
import fs from "fs";
import { google } from "googleapis";
import http from "http";
import open from "open";
import path from "path";
import puppeteer from "puppeteer";
import { Server } from "socket.io";
import { fileURLToPath } from "url";
import xlsx from "xlsx";
import logger from "./logger.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config();

const app = express();
app.use(
  cors({
    origin: "http://localhost:3001", // Frontend URL
    methods: ["GET", "POST"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);

const server = http.createServer(app);
const io = new Server(server, {
  cors: {
    origin: "*",
    methods: ["GET", "POST"],
  },
});

// Serve static files from the public directory
app.use(express.static(path.join(__dirname, "public")));

const port = 3000;

const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REDIRECT_URI = `http://localhost:${port}/oauth2callback`;
const TOKEN_PATH = "token.json";

const oauth2Client = new google.auth.OAuth2(
  CLIENT_ID,
  CLIENT_SECRET,
  REDIRECT_URI
);
const SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"];

let verificationCode = "";

const getDataFromSheet = () => {
  try {
    const workbook = xlsx.readFile("testing-1.xlsx");
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const sheetData = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    console.log("Sheet Data:", sheetData);

    if (!Array.isArray(sheetData) || sheetData.length === 0) {
      console.error("No data found in the sheet or invalid format");
      return [];
    }

    const hashMap = {};
    const hotels = [];
    let cnt = 0;

    for (const item of sheetData) {
      if (!item["Property ID"]) {
        console.warn("Skipping row with no Property ID:", item);
        continue;
      }

      if (!hashMap[item["Property ID"]]) {
        hashMap[item["Property ID"]] = ++cnt;
        hotels.push({
          id: item["Property ID"],
          idList: [item["Reservation ID"]],
        });
      } else {
        hotels[hashMap[item["Property ID"]] - 1].idList.push(
          item["Reservation ID"]
        );
      }
    }

    console.log("Processed Hotels:", hotels);
    return hotels;
  } catch (error) {
    console.error("Error reading sheet:", error);
    return [];
  }
};

// Add this helper function at the top level
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Load token if it exists
function loadToken() {
  if (fs.existsSync(TOKEN_PATH)) {
    const token = JSON.parse(fs.readFileSync(TOKEN_PATH));
    oauth2Client.setCredentials(token);
    return true;
  }
  return false;
}

// Fetch verification code from Gmail
async function getVerificationCode() {
  try {
    const gmail = google.gmail({ version: "v1", auth: oauth2Client });
    const res = await gmail.users.messages.list({
      userId: "me",
      maxResults: 5,
    });

    if (!res.data.messages) {
      logger.info("No new emails found.");
      return null;
    }

    for (const msg of res.data.messages) {
      const email = await gmail.users.messages.get({
        userId: "me",
        id: msg.id,
      });
      const body = email.data.snippet;
      logger.info("Email body:", body);
      const codeMatch = body.match(/\b\d{6,10}\b/);
      logger.info("Code match:", codeMatch);

      if (codeMatch) {
        return codeMatch[0];
      }
    }

    logger.info("No verification code found in recent emails.");
    return null;
  } catch (error) {
    logger.error("Error fetching emails:", error.message);
    return null;
  }
}

// Utility function for delays
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const randomDelay = () => Math.floor(Math.random() * (3000 - 1000) + 1000);

// Puppeteer Login Function
async function loginToExpediaPartner(
  email = process.env.EMAIL,
  password = process.env.PASSWORD
) {
  let browser = null;
  try {
    browser = await puppeteer.launch({
      headless: false,
      defaultViewport: null,
      args: [
        "--start-maximized",
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-web-security",
        "--disable-features=IsolateOrigins,site-per-process",
      ],
      timeout: 60000,
    });

    const page = await browser.newPage();
    await page.setDefaultNavigationTimeout(60000);
    await page.setDefaultTimeout(60000);

    // Navigate to partner central
    logger.info("Navigating to Expedia Partner Central...");
    await page.goto(
      "https://www.expediapartnercentral.com/Account/Logon?signedOff=true",
      {
        waitUntil: ["networkidle0", "domcontentloaded"],
        timeout: 60000,
      }
    );

    logger.info("Waiting for page load...");

    await delay(randomDelay());

    await page.evaluate(() => {
      window.scrollBy(0, 200); // Scroll down by 200 pixels
    });
    // Wait for email input
    await page.waitForSelector("#emailControl");

    // Type email slowly, character by character
    for (let char of email) {
      await page.type("#emailControl", char, { delay: 100 }); // 100ms delay between each character
    }

    // Click continue button
    await page.click("#continueButton");

    // Wait before entering password
    logger.info("Waiting for password page to load...");

    // Wait for password page to be fully loaded
    try {
      logger.info("Waiting for password page to fully load...");

      // Try to find the password input field with a try-catch to handle both possible selectors
      let passwordInputFound = false;

      try {
        // First try to find #password-input
        const passwordInput = await page.waitForSelector("#password-input", {
          visible: true,
          timeout: 15000, // Shorter timeout for first attempt
        });

        if (passwordInput) {
          passwordInputFound = true;

          // Add a significant delay to ensure the page is fully loaded and stable
          await delay(3000);

          // Verify the password field is actually ready for input
          const isInputReady = await page.evaluate(() => {
            const input = document.querySelector("#password-input");
            return input && !input.disabled && document.activeElement !== input;
          });

          if (!isInputReady) {
            logger.info("Password input not fully ready, waiting longer...");
            await delay(2000);
          }

          // Click on the password field first to ensure focus
          await page.click("#password-input");
          await delay(1000);

          // Clear the field in case there's any text
          await page.evaluate(() => {
            document.querySelector("#password-input").value = "";
          });
          await delay(500);

          logger.info("Password page fully loaded, entering password...");

          // Type password slowly with increased delays
          for (let char of password) {
            await page.type("#password-input", char, { delay: 150 }); // Increased delay
            await delay(100); // Increased delay between characters
          }

          // Wait longer before clicking submit to ensure password is fully entered
          logger.info("Password entered, waiting before clicking submit...");
          await delay(5000);

          // Verify password was entered correctly
          const enteredPassword = await page.evaluate(() => {
            return document.querySelector("#password-input").value;
          });

          if (enteredPassword.length !== password.length) {
            logger.warn(
              `Password entry issue: expected ${password.length} chars but got ${enteredPassword.length}`
            );

            // Re-enter password if needed
            await page.evaluate(() => {
              document.querySelector("#password-input").value = "";
            });
            await delay(1000);

            // Try again with even slower typing
            for (let char of password) {
              await page.type("#password-input", char, { delay: 200 });
              await delay(150);
            }
            await delay(2000);
          }

          // Click the login button
          logger.info("Clicking password continue button...");
          await page.click("#password-continue");
        }
      } catch (error) {
        logger.info(
          "Could not find #password-input, trying #passwordControl instead:",
          error.message
        );
        passwordInputFound = false;
      }

      // If #password-input wasn't found, try #passwordControl
      if (!passwordInputFound) {
        try {
          // Check if #passwordControl exists
          const passwordControlExists = await page.evaluate(() => {
            return !!document.querySelector("#passwordControl");
          });

          if (!passwordControlExists) {
            logger.info(
              "Neither #password-input nor #passwordControl found. Checking page content..."
            );
            const pageContent = await page.content();
            logger.info("Page title: " + (await page.title()));
            throw new Error("Password input field not found on the page");
          }

          // Add a significant delay to ensure the page is fully loaded and stable
          await delay(3000);

          // Verify the password field is actually ready for input
          const isInputReady = await page.evaluate(() => {
            const input = document.querySelector("#passwordControl");
            return input && !input.disabled && document.activeElement !== input;
          });

          if (!isInputReady) {
            logger.info("Password input not fully ready, waiting longer...");
            await delay(2000);
          }

          // Click on the password field first to ensure focus
          await page.click("#passwordControl");
          await delay(1000);

          // Clear the field in case there's any text
          await page.evaluate(() => {
            document.querySelector("#passwordControl").value = "";
          });
          await delay(500);

          logger.info("Password page fully loaded, entering password...");

          // Type password slowly with increased delays
          for (let char of password) {
            await page.type("#passwordControl", char, { delay: 150 }); // Increased delay
            await delay(100); // Increased delay between characters
          }

          // Wait longer before clicking submit to ensure password is fully entered
          logger.info("Password entered, waiting before clicking submit...");
          await delay(5000);

          // Verify password was entered correctly
          const enteredPassword = await page.evaluate(() => {
            return document.querySelector("#passwordControl").value;
          });

          if (enteredPassword.length !== password.length) {
            logger.warn(
              `Password entry issue: expected ${password.length} chars but got ${enteredPassword.length}`
            );

            // Re-enter password if needed
            await page.evaluate(() => {
              document.querySelector("#passwordControl").value = "";
            });
            await delay(1000);

            // Try again with even slower typing
            for (let char of password) {
              await page.type("#passwordControl", char, { delay: 200 });
              await delay(150);
            }
            await delay(2000);
          }

          // Click the login button
          logger.info("Clicking password continue button...");
          await page.click("#signInButton");
        } catch (error) {
          logger.error("Error handling password input:", error.message);
          throw error;
        }
      }
    } catch (error) {
      logger.info("Error during password entry:", error.message);
      throw error;
    }

    // Wait for verification code page using the correct selector
    logger.info("Waiting for verification page...");
    await page.waitForSelector('input[name="passcode-input"]', {
      visible: true,
      timeout: 60000,
    });

    // Add delay before fetching verification code
    logger.info("Waiting for verification email...");
    await delay(15000); // Wait 15 seconds for email to arrive

    // Get verification code
    const code = await getVerificationCode();
    if (!code) {
      throw new Error("Failed to get verification code from email");
    }
    logger.info("Got verification code:", code);

    // Enter verification code using the correct selector
    await page.type('input[name="passcode-input"]', code, { delay: 100 });
    await delay(randomDelay());

    // await verifyButton.click()
    const verifyButtonHandle = await page.$(
      'button[data-testid="passcode-submit-button"]'
    );

    if (!verifyButtonHandle) {
      throw new Error("Verify button not found");
    }

    // Check if the button is disabled
    const isDisabled = await page.evaluate(
      (button) => button.disabled,
      verifyButtonHandle
    );

    if (isDisabled) {
      throw new Error("Verify button is disabled");
    }

    // Click the button
    await verifyButtonHandle.click();
    logger.info("Clicked the verify button successfully!");

    // Wait for successful login
    await page.waitForNavigation({
      waitUntil: "networkidle0",
      timeout: 60000,
    });

    logger.info("Login successful!");

    const sheetData = getDataFromSheet();
    const allReservations = [];

    for (const item of sheetData) {
      const propertyName = item.id;

      if (propertyName) {
        // Wait for property table to load
        await page.waitForSelector(".fds-data-table-wrapper", {
          visible: true,
          timeout: 30000,
        });

        // Wait for property search input
        await page.waitForSelector(
          ".all-properties__search input.fds-field-input"
        );

        // Get property ID from query params
        logger.info(`Searching for property ID: ${propertyName}`);

        // Type property ID in search
        await page.type(
          ".all-properties__search input.fds-field-input",
          String(propertyName),
          { delay: 500 }
        );

        // Wait for search results
        await delay(2000);

        // Find and click the property link with more specific selector
        try {
          // Wait for search results to update
          await page.waitForSelector("tbody tr", {
            visible: true,
            timeout: 10000,
          });

          // Find and click the property link
          const clicked = await page.evaluate((searchId) => {
            const rows = Array.from(document.querySelectorAll('tbody tr'));
            for (const row of rows) {
              const idElement = row.querySelector('.property-cell__property-id span');
              if (idElement && idElement.textContent.includes(searchId)) {
                const link = row.querySelector('.property-cell__property-name a');
                if (link) {
                  link.click();
                  return true;
                }
              }
            }
            return false;
          }, String(propertyName));

          if (clicked) {
            logger.info(`Found and clicked property with ID: ${propertyName}`);
            
            // Wait for navigation
            await Promise.all([
              page.waitForNavigation({
                waitUntil: "networkidle0",
                timeout: 30000,
              }),
              delay(8000),
            ]);

            logger.info("Successfully navigated to property page");
          } else {
            throw new Error(`Could not find property with ID: ${propertyName}`);
          }
        } catch (error) {
          logger.error(`Error finding/clicking property: ${error.message}`);
          throw error;
        }
      }

      logger.info("Looking for Reservations link...");

      try {
        // Wait for the drawer content to load
        await page.waitForSelector(".uitk-drawer-content", {
          visible: true,
          timeout: 30000,
        });

        // Click using JavaScript with the exact structure
        const clicked = await page.evaluate(() => {
          const reservationsItem = Array.from(
            document.querySelectorAll(".uitk-action-list-item-content")
          ).find((item) => {
            const textDiv = item.querySelector(".uitk-text.overflow-wrap");
            return textDiv && textDiv.textContent.trim() === "Reservations";
          });

          if (reservationsItem) {
            const link = reservationsItem.querySelector(
              "a.uitk-action-list-item-link"
            );
            if (link) {
              link.click();
              return true;
            }
          }
          return false;
        });

        if (!clicked) {
          throw new Error("Could not find or click Reservations link");
        }

        // Wait for navigation to complete
        await Promise.all([
          page.waitForNavigation({
            waitUntil: "networkidle0",
            timeout: 80000,
          }),
          delay(8000),
        ]);

        logger.info("Successfully navigated to Reservations page");

        // Wait for date filters to be visible
        logger.info("Waiting for date filters...");
        await page.waitForSelector(
          'input[type="radio"][name="dateTypeFilter"]',
          {
            visible: true,
            timeout: 80000,
          }
        );

        // Get the current URL
        const currentUrl = page.url();
        console.log(`Current tab URL: ${currentUrl}`);

        for (const chunk of item.idList) {
          logger.info(`Processing id: ${chunk}`);

          const chunkReservations = await processReservationsPage(page, chunk, propertyName, propertyName, browser);
          allReservations.push(...chunkReservations);
        }

        logger.info(
          `Found total ${allReservations.length} reservations across all chunks`
        );
        await delay(5000);
        try {
          // Wait for the header to be visible
          await page.waitForSelector("header.tpg-navigation__header", {
            visible: true,
            timeout: 5000,
          });

          // Try multiple approaches to click the logo
          const clicked = await page.evaluate(() => {
            // Try finding the logo link
            const logoLink = document.querySelector('header.tpg-navigation__header a.tpg-navigation__logo_container');
            if (logoLink) {
              logoLink.click();
              return true;
            }
            return false;
          });

          if (!clicked) {
            // If direct click failed, try using the href
            const href = await page.evaluate(() => {
              const logoLink = document.querySelector('header.tpg-navigation__header a.tpg-navigation__logo_container');
              return logoLink ? logoLink.href : null;
            });

            if (href) {
              await page.goto(href, { waitUntil: 'networkidle0' });
            } else {
              throw new Error('Could not find navigation logo link');
            }
          }

          await delay(2000); // Wait for navigation
        } catch (error) {
          logger.warn("Could not click navigation logo:", error.message);
          // Try alternative navigation
          try {
            await page.goto('https://apps.expediapartnercentral.com/', { waitUntil: 'networkidle0' });
            await delay(2000);
          } catch (navError) {
            logger.error("Failed to navigate to home page:", navError.message);
          }
        }

        // Add 5 second delay before processing next property
        logger.info("Waiting 5 seconds before processing next property...");
        await delay(5000);
      } catch (error) {
        logger.error("Error finding/clicking Reservations:", error.message);
        throw error;
      }
    }

    // Export all data to Excel after processing all properties
    if (allReservations.length > 0) {
      // Get current date and time for filename
      const now = new Date();
      const timestamp = now.toISOString().replace(/[:.]/g, "-");

      // Save all reservations to Excel with timestamp
      const workbook = xlsx.utils.book_new();
      const wsData = [
        [
          "Property ID",
          "Property Name",
          "Guest Name",
          "Reservation ID",
          "Confirmation Code",
          "Check-in Date",
          "Check-out Date",
          "Room Type",
          "Booking Amount",
          "Booked Date",
          "Card Number",
          "Expiry Date",
          "CVV",
          "Has Card Info",
          "Has Payment Info",
          "Total Guest Payment",
          "Cancellation Fee",
          "Expedia Compensation",
          "Total Payout",
          "Details",
          "Status",
          "Amount to charge/refund",
        ],
        ...allReservations.map((res) => [
          res.propertyId || "N/A",
          res.propertyName || "N/A",
          res.guestName,
          res.reservationId,
          res.confirmationCode,
          res.checkInDate,
          res.checkOutDate,
          res.roomType,
          res.bookingAmount,
          res.bookedDate,
          res.cardNumber || "N/A",
          res.expiryDate || "N/A",
          res.cvv || "N/A",
          res.hasCardInfo ? "Yes" : "No",
          res.hasPaymentInfo ? "Yes" : "No",
          res.totalGuestPayment || "N/A",
          res.cancellationFee || "N/A",
          res.expediaCompensation || "N/A",
          res.totalPayout || "N/A",
          res.amountToChargeOrRefund || "N/A",
          res.status || "Active",
          res.amount || "N/A",
        ]),
      ];

      const ws = xlsx.utils.aoa_to_sheet(wsData);
      xlsx.utils.book_append_sheet(workbook, ws, "Reservations");
      xlsx.writeFile(workbook, `reservations_${timestamp}.xlsx`);
      logger.info(`Saved reservation data to reservations_${timestamp}.xlsx`);

      // Log summary of processed reservations
      logger.info(`Total reservations exported: ${allReservations.length}`);
      const processedDates = new Set(allReservations.map((r) => r.checkInDate));
      if (processedDates.size === 0) {
        logger.warn(
          "Warning: No dates were successfully processed in this export"
        );
      }
    }
  } catch (error) {
    logger.error(`Error finding/clicking property: ${error.message}`);
    if (browser) await browser.close();
    throw error;
  }
}

// New function to process reservations on a single page
async function processReservationsPage(page, id, propertyId, propertyName, browser) {
  try {
    try {
      // Wait for the page to be fully loaded
      await page.waitForSelector(".fds-layout", {
        visible: true,
        timeout: 30000,
      });

      // Try to find the search input using multiple possible selectors
      const searchInputSelectors = [
        'input[name="searchInput"]',
        "input.fds-field-input",
        'input[type="text"]',
        ".fds-field-input",
      ];

      let searchInput = null;
      for (const selector of searchInputSelectors) {
        try {
          searchInput = await page.waitForSelector(selector, {
            visible: true,
            timeout: 5000,
          });
          if (searchInput) break;
        } catch (e) {
          continue;
        }
      }

      if (!searchInput) {
        throw new Error("Could not find search input field");
      }

      // Click the input field first
      await searchInput.click();
      await delay(1000);

      // Clear any existing value
      await page.evaluate(() => {
        const input =
          document.querySelector('input[name="searchInput"]') ||
          document.querySelector("input.fds-field-input") ||
          document.querySelector('input[type="text"]');
        if (input) input.value = "";
      });

      // Type the ID into the search input - convert id to string and type it directly
      const idString = String(id);
      await page.type('input[name="searchInput"]', idString, { delay: 150 });

      // Wait for the save button to be visible and clickable
      await page.waitForSelector("#save-button", {
        visible: true,
        timeout: 10000,
      });

      // Click the save button
      await page.click("#save-button");

      // Wait for the search to complete
      await delay(2000);

      // Final verification
      const finalCount = await page.evaluate(() => {
        return document.querySelectorAll("td.guestName button.guestNameLink")
          .length;
      });

      logger.info(`Final reservation count: ${finalCount}`);

      if (finalCount === 0) {
        logger.info("No reservations found after multiple attempts");
        return [];
      }

      // After date range is applied and before scraping data
      // logger.info('Setting results per page to 100...')
      // await page.waitForSelector('.fds-pagination-selector select')
      // await page.click('.fds-pagination-selector select')
      // await page.select('.fds-pagination-selector select', '100')

      // Wait for data to reload with 100 records
      await delay(3000);
      await page.waitForSelector("table.fds-data-table tbody tr", {
        visible: true,
        timeout: 30000,
      });

      // Initialize array for all reservations with Set for tracking duplicates
      const pageReservations = [];
      const processedReservationIds = new Set();

      // Function to check if there's a next page
      const hasNextPage = async () => {
        return await page.evaluate(() => {
          const nextButton = document.querySelector(
            ".fds-pagination-button.next button"
          );
          return nextButton && !nextButton.disabled;
        });
      };

      // Function to get total results count
      const getTotalResults = async () => {
        const resultsText = await page.$eval(
          ".fds-pagination-showing-result",
          (el) => el.textContent
        );
        const match = resultsText.match(/of (\d+) Results/);
        return match ? parseInt(match[1]) : 0;
      };

      const totalResults = await getTotalResults();
      logger.info(`Total reservations to fetch: ${totalResults}`);

      let currentPage = 1;
      let hasMore = true;

      while (hasMore) {
        try {
          logger.info(`Processing page ${currentPage}...`);

          // Wait for table data to load
          await page.waitForSelector("table.fds-data-table tbody tr", {
            visible: true,
            timeout: 30000,
          });
          await delay(5000);

          // Get reservations from current page
          const rows = await page.$$("table.fds-data-table tbody tr");

          for (const row of rows) {
            try {
              // Get basic data first
              const basicData = await page.evaluate((row) => {
                return {
                  guestName:
                    row
                      .querySelector(
                        "td.guestName button.guestNameLink span.fds-button2-label"
                      )
                      ?.textContent.trim() || "",
                  reservationId:
                    row
                      .querySelector("td.reservationId div.fds-cell")
                      ?.textContent.trim() || "",
                  confirmationCode:
                    row
                      .querySelector(
                        "td.confirmationCode label.confirmationCodeLabel"
                      )
                      ?.textContent.trim() || "",
                  checkInDate:
                    row.querySelector("td.checkInDate")?.textContent.trim() ||
                    "",
                  checkOutDate:
                    row.querySelector("td.checkOutDate")?.textContent.trim() ||
                    "",
                  roomType:
                    row.querySelector("td.roomType")?.textContent.trim() || "",
                  bookingAmount:
                    row
                      .querySelector("td.bookingAmount .fds-currency-value")
                      ?.textContent.trim() || "",
                  bookedDate:
                    row.querySelector("td.bookedOnDate")?.textContent.trim() ||
                    "",
                };
              }, row);

              // Check if we've already processed this reservation
              if (processedReservationIds.has(basicData.reservationId)) {
                logger.info(
                  `Skipping duplicate reservation: ${basicData.reservationId}`
                );
                continue;
              }

              // Add to processed set
              processedReservationIds.add(basicData.reservationId);

              // Get card details
              const guestNameButton = await row.$(
                "td.guestName button.guestNameLink"
              );
              await guestNameButton.click();

              // Wait for initial dialog to appear with timeout
              try {
                await Promise.race([
                  page.waitForSelector(".fds-dialog", {
                    visible: true,
                    timeout: 8000,
                  }),
                  new Promise((_, reject) =>
                    setTimeout(() => reject(new Error("Dialog timeout")), 8000)
                  ),
                ]);
              } catch (error) {
                logger.info(
                  "Dialog did not appear within timeout, skipping to next reservation"
                );
                continue;
              }

              // Wait a bit for content to load
              await delay(2000);

              // Scroll to the bottom of dialog content and wait
              await page.evaluate(() => {
                const dialogContent = document.querySelector(
                  ".fds-dialog-content"
                );
                if (dialogContent) {
                  dialogContent.scrollTo(0, dialogContent.scrollHeight);
                }
              });

              // Wait for content to load after scroll
              await delay(2000);

              // Look for the "See card activity" button and click it in a new tab
              let remainingBalance = "N/A";
              try {
                const seeCardActivityButton = await page.$('.fds-cell.all-y-gutter-16 button.fds-button2.utility.small');
                
                if (seeCardActivityButton) {
                  logger.info("Found 'See card activity' button, clicking it in a new tab...");
                  
                  // Get href or onclick URL from the button
                  const buttonUrl = await page.evaluate(() => {
                    const button = document.querySelector('.fds-cell.all-y-gutter-16 button.fds-button2.utility.small');
                    if (!button) return null;
                    
                    // Click the button but prevent navigation by returning the URL
                    const originalOpen = window.open;
                    let capturedUrl = null;
                    
                    // Override window.open temporarily to capture the URL
                    window.open = (url) => {
                      capturedUrl = url;
                      return { focus: () => {} }; // Mock window object
                    };
                    
                    // Simulate click to trigger any onclick handlers
                    button.click();
                    
                    // Restore original window.open
                    window.open = originalOpen;
                    
                    return capturedUrl;
                  });
                  
                  if (buttonUrl) {
                    logger.info(`Opening card activity URL in new tab: ${buttonUrl}`);
                    
                    // Create a new page/tab
                    const newPage = await browser.newPage();
                    await newPage.goto(buttonUrl, { waitUntil: 'networkidle0', timeout: 30000 });
                    
                    logger.info("New tab opened for card activity");
                    await delay(5000); // Give more time for the page to fully load
                    
                    // Scrape the remaining balance
                    remainingBalance = await newPage.evaluate(() => {
                      // Try multiple selectors to find the remaining balance
                      const selectors = [
                        '.evc-mock-card-remaining-balance .fds-currency-value',
                        '.remaining-balance .fds-currency-value',
                        '[class*="remaining-balance"] .fds-currency-value',
                        '[class*="balance"] .fds-currency-value',
                        '.fds-currency-value'
                      ];
                      
                      for (const selector of selectors) {
                        const elements = document.querySelectorAll(selector);
                        for (const element of elements) {
                          // Check if parent contains text about balance
                          const parent = element.closest('div');
                          if (parent && parent.textContent.toLowerCase().includes('balance')) {
                            return element.textContent.trim();
                          }
                        }
                      }
                      
                      // If we couldn't find a specific balance element, try to get any currency value
                      const anyBalance = document.querySelector('.fds-currency-value');
                      return anyBalance ? anyBalance.textContent.trim() : "N/A";
                    });
                    
                    logger.info(`Scraped remaining balance: ${remainingBalance}`);
                    
                    // Take screenshot for debugging if needed
                    await newPage.screenshot({ path: 'card-activity.png' });
                    
                    // Close the new tab
                    await newPage.close();
                    logger.info("Closed card activity tab");
                  } else {
                    logger.info("Could not capture URL from 'See card activity' button, skipping");
                  }
                } else {
                  logger.info("'See card activity' button not found, skipping");
                }
              } catch (error) {
                logger.warn(`Error processing card activity: ${error.message}`);
              }

              // Get card details with retry mechanism
              let cardData = null;
              let paymentData = null;
              let remainingAmountToCharge = null;
              let amountToRefund = null;
              let status = "None"; // Default status
              let additionalText = ""; // New variable to store additional text
              let retries = 0;
              while (retries < 3) {
                try {
                  // First check for evcCardBase element
                  const hasEvcCard = await page.evaluate(() => {
                    const evcCardBase = document.querySelector('.evcCardBase');
                    if (evcCardBase) {
                      // Get status badge if it exists
                      const statusBadge = evcCardBase.querySelector('.fds-grid.statusBadge .fds-badge');
                      return {
                        exists: true,
                        status: statusBadge ? statusBadge.textContent.trim() : 'None'
                      };
                    }
                    return { exists: false };
                  });

                  if (hasEvcCard.exists) {
                    status = hasEvcCard.status;
                    // Get card details from evcCardBase
                    cardData = await page.evaluate(() => {
                      const cardNumber = document.querySelector('.evcCardBase .cardNumber.replay-conceal bdi')?.textContent.trim() || '';
                      const expiryDate = document.querySelector('.evcCardBase .cardDetails .fds-cell.all-cell-1-4.fds-type-color-primary.replay-conceal')?.textContent.trim() || '';
                      const cvv = document.querySelectorAll('.evcCardBase .cardDetails .fds-cell.all-cell-1-4.fds-type-color-primary.replay-conceal')[1]?.textContent.trim() || '';

                      // Get additional text information
                      const additionalTextElements = Array.from(document.querySelectorAll('.fds-cell.all-y-gutter-12 div, .fds-cell.sidePanelSection, .fds-cell.fds-type-color-attention.fds-grid .fds-cell.all-cell-fill'));
                      const additionalText = additionalTextElements
                        .map(el => el.textContent.trim())
                        .filter(text => text && 
                          !text.includes('See card activity') && 
                          !text.includes('contact us') &&
                          !text.includes('Show contact details'))
                        .join(' | ');

                      if (cardNumber) {
                        return {
                          cardNumber,
                          expiryDate,
                          cvv,
                          status: status,
                          additionalText
                        };
                      }
                      return null;
                    });
                  }

                  // Always try to get payment information regardless of card data
                  paymentData = await page.evaluate(() => {
                    // Find all payment summary sections
                    const paymentSummary = document.querySelector('.fds-card-content');
                    if (!paymentSummary) return null;

                    // Helper function to find value by section title
                    const findValueByTitle = (titleText) => {
                      const sections = Array.from(paymentSummary.querySelectorAll('.fds-grid'));
                      for (const section of sections) {
                        const title = section.querySelector('.sidePanelSectionTitle');
                        if (title && title.textContent.trim() === titleText) {
                          const value = section.querySelector('.fds-currency-value');
                          return value ? value.textContent.trim() : '';
                        }
                      }
                      return '';
                    };

                    // Get all payment values
                    const cancellationFee = findValueByTitle('Cancellation fee');
                    const expediaCompensation = findValueByTitle('Expedia compensation');
                    const totalPayout = findValueByTitle('Your total payout');
                    const totalGuestPayment = findValueByTitle('Total guest payment');

                    if (cancellationFee || expediaCompensation || totalPayout) {
                      return {
                        totalGuestPayment,
                        cancellationFee,
                        expediaCompensation,
                        totalPayout,
                      };
                    }
                    return null;
                  });

                  // Extract "Remaining amount to charge" and "Amount to refund"
                  const additionalPaymentInfo = await page.evaluate(() => {
                    // Find "Remaining amount to charge"
                    const remainingAmountSection = Array.from(
                      document.querySelectorAll(".fds-cell.sidePanelSection")
                    ).find((section) =>
                      section.textContent.includes("Remaining amount to charge")
                    );

                    const remainingAmount =
                      remainingAmountSection
                        ?.querySelector(".fds-currency-value")
                        ?.textContent.trim() || "";

                    // Find "Amount to refund"
                    const refundSection = Array.from(
                      document.querySelectorAll(".fds-grid.sidePanelSection")
                    ).find((section) =>
                      section.textContent.includes("Amount to refund")
                    );

                    const refundAmount =
                      refundSection
                        ?.querySelector(".fds-currency-value")
                        ?.textContent.trim() || "";

                    return {
                      remainingAmountToCharge: remainingAmount,
                      amountToRefund: refundAmount,
                    };
                  });

                  if (additionalPaymentInfo) {
                    remainingAmountToCharge =
                      additionalPaymentInfo.remainingAmountToCharge;
                    amountToRefund = additionalPaymentInfo.amountToRefund;

                    if (remainingAmountToCharge) {
                      logger.info(
                        `Found Remaining amount to charge: ${remainingAmountToCharge}`
                      );
                    }

                    if (amountToRefund) {
                      logger.info(`Found Amount to refund: ${amountToRefund}`);
                    }
                  }

                  // Break the loop if we got either card data or payment data
                  if (cardData || paymentData) {
                    break;
                  }

                  retries++;
                  await delay(1000);
                } catch (e) {
                  retries++;
                  await delay(1000);
                }
              }

              //////////////////////////////////////////////////////////////
              //close the side panel
              //////////////////////////////////////////////////////////////
              try {
                await page.click(".fds-dialog-header button.dialog-close");
                await delay(1500);
              } catch (e) {
                logger.warn("Warning: Could not close dialog normally");
              }

              // Get property name from the header with more specific selector
              const propertyInfo = await page.evaluate(() => {
                // Try multiple selectors to find the property name
                const selectors = [
                  '.tpg-navigation__header__dropdown-property-details .fds-dropdown-button-label',
                  '.tpg-navigation__header__dropdown-property-details button span',
                  '.tpg-navigation__header__dropdown-property-details',
                  '.tpg-navigation__header__dropdown-property-details .fds-button2-label'
                ];
                
                for (const selector of selectors) {
                  const element = document.querySelector(selector);
                  if (element) {
                    const text = element.textContent.trim();
                    if (text && text !== '') {
                      return text;
                    }
                  }
                }
                return '';
              });

              // When adding to pageReservations array, include property info
              pageReservations.push({
                ...basicData,
                ...(cardData || {}),
                ...(paymentData || {}),
                propertyId: propertyId,
                propertyName: propertyInfo || propertyName, // Use propertyName as fallback
                hasCardInfo: !!cardData,
                hasPaymentInfo: !!paymentData,
                remainingAmountToCharge: remainingAmountToCharge || "N/A",
                amountToRefund: amountToRefund || "N/A",
                amountToChargeOrRefund: cardData?.additionalText || remainingAmountToCharge || amountToRefund || "N/A",
                status: status,
                amount: remainingBalance,
              });
            } catch (error) {
              logger.info(`Error processing reservation: ${error.message}`);
              if (basicData) {
                pageReservations.push({
                  ...basicData,
                  cardNumber: "N/A",
                  expiryDate: "N/A",
                  cvv: "N/A",
                  remainingAmountToCharge: "N/A",
                  amountToRefund: "N/A",
                  amountToChargeOrRefund: "N/A",
                });
              }
            }
          }

          logger.info(
            `Processed ${pageReservations.length} of ${totalResults} reservations`
          );

          // Check if there's a next page
          hasMore = await hasNextPage();
          if (hasMore) {
            // Scroll down smoothly before clicking next page
            await page.evaluate(() => {
              window.scrollBy({
                top: 300,
                behavior: "smooth",
              });
            });
            await delay(1500); // Wait for scroll animation

            await page.click(".fds-pagination-button.next button");
            await delay(2000);
            currentPage++;
          }
        } catch (pageError) {
          logger.info(
            `Error processing page ${currentPage}: ${pageError.message}`
          );
          // Try to recover by reloading the page
          await page.reload({ waitUntil: "networkidle0" });
          await delay(5000);
        }
      }

      logger.info(
        `Date scraping completed. Found total ${pageReservations.length} reservations on this tab`
      );

      // Log if no reservations were found for this date range
      if (pageReservations.length === 0) {
        logger.warn(
          `No reservations found for date range. This date range may be missing data.`
        );
      }
      return pageReservations;
    } catch (error) {
      logger.error(`Error processing tab: ${error.message}`);
      return [];
    }
  } catch (error) {
    logger.error(`Error processing tab: ${error.message}`);
    return [];
  }
}

// API endpoint to get logs
app.get("/api/data", (req, res) => {
  try {
    const data = JSON.parse(
      fs.readFileSync(path.join(__dirname, "data.json"), "utf8")
    );
    
    // Filter logs from the last hour
    const oneHourAgo = new Date(Date.now() - 60 * 60 * 1000);
    const filteredLogs = data.logs.filter(log => {
      return new Date(log.timestamp) > oneHourAgo;
    });
    
    res.json({ logs: filteredLogs });
  } catch (error) {
    res.status(500).json({ error: "Error reading logs" });
  }
});

// Function to cleanup old logs (older than 1 hour)
function cleanupOldLogs() {
  try {
    const filePath = path.join(__dirname, "data.json");
    
    // Read current data
    const data = JSON.parse(fs.readFileSync(filePath, "utf8"));
    
    // Filter logs from the last hour
    const oneHourAgo = new Date(Date.now() - 60 * 60 * 1000);
    const filteredLogs = data.logs.filter(log => {
      return new Date(log.timestamp) > oneHourAgo;
    });
    
    // Write filtered data back to file
    fs.writeFileSync(filePath, JSON.stringify({ logs: filteredLogs }, null, 2));
    
    logger.info(`Cleaned up logs: Removed ${data.logs.length - filteredLogs.length} old entries`);
  } catch (error) {
    logger.error(`Error cleaning up logs: ${error.message}`);
  }
}

// Run cleanup every 15 minutes
setInterval(cleanupOldLogs, 15 * 60 * 1000);

// Watch for JSON file changes
fs.watch(path.join(__dirname, "data.json"), () => {
  try {
    const data = JSON.parse(
      fs.readFileSync(path.join(__dirname, "data.json"), "utf8")
    );
    io.emit("update", data); // Broadcast updates
  } catch (error) {
    console.error("Error reading logs:", error);
  }
});

// WebSocket connection handling
io.on("connection", (socket) => {
  console.log("Client connected");

  fs.readFile(path.join(__dirname, "data.json"), "utf8", (err, data) => {
    if (!err) socket.emit("update", JSON.parse(data)); // Send initial data
  });

  socket.on("disconnect", () => console.log("Client disconnected"));
});

// Express routes
app.get("/auth", async (req, res) => {
  const authUrl = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: SCOPES,
  });

  res.redirect(authUrl);
});

app.get("/oauth2callback", async (req, res) => {
  const code = req.query.code;
  if (!code) {
    return res.status(400).send("Authorization code not found.");
  }

  try {
    const { tokens } = await oauth2Client.getToken(code);
    oauth2Client.setCredentials(tokens);
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
    // res.send('Authentication successful! You can close this window.')
    res.redirect(process.env.FRONTEND_REDIRECT_URI);
  } catch (error) {
    res.status(500).send("Error retrieving access token: " + error.message);
  }
});
// Independent API endpoint for Expedia login automation
app.get("/api/expedia", async (req, res) => {
  const { email, password, propertyName } = req.query;

  if (!email || !password) {
    return res.status(400).json({
      success: false,
      message: "Email, password, and are required",
    });
  }

  try {
    if (!loadToken()) {
      return res
        .status(401)
        .json({ success: false, message: "Gmail authentication required" });
    }

    // Call loginToExpediaPartner with the validated dates
    await loginToExpediaPartner(email, password, propertyName);

    res.json({
      success: true,
      message: "Successfully processed",
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// Start the Express server
server.listen(port, () => {
  logger.info(`Server running at http://localhost:${port}`);
  
  // Run cleanup immediately when server starts
  cleanupOldLogs();
  
  if (!loadToken()) {
    logger.info("Opening browser for authentication...");
    open(`http://localhost:${port}/auth`);
  }
});