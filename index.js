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
    const workbook = xlsx.readFile("VNP_sheet.xlsx");
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
      if (!item["Property Name"]) {
        console.warn("Skipping row with no Property Name:", item);
        continue;
      }

      if (!hashMap[item["Property Name"]]) {
        hashMap[item["Property Name"]] = ++cnt;
        hotels.push({
          name: item["Property Name"],
          idList: [item["Reservation ID"]],
        });
      } else {
        hotels[hashMap[item["Property Name"]] - 1].idList.push(
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
      const propertyName = item.name;

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

        // Get property name from query params
        logger.info(`Searching for property: ${propertyName}`);

        // Type property name in search
        await page.type(
          ".all-properties__search input.fds-field-input",
          propertyName,
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

          // More specific selector for the property link
          const propertySelector = `.property-cell__property-name`;

          const propertyLink = await page.waitForSelector(propertySelector, {
            visible: true,
            timeout: 10000,
          });

          if (propertyLink) {
            // Get the text to verify it's the right property
            const linkText = await page.evaluate(
              (el) => el.textContent,
              propertyLink
            );
            logger.info(`Found property: ${linkText}, clicking...`);

            try {
              // Click the link and wait for navigation
              await Promise.all([
                page.waitForNavigation({
                  waitUntil: "networkidle0",
                  timeout: 30000,
                }),
                propertyLink.click(),
              ]);

              // Wait for the new page to load
              await delay(8000);
            } catch (error) {
              console.error(error.message);
            }
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

          const chunkReservations = await processReservationsPage(page, chunk);
          allReservations.push(...chunkReservations);
        }

        logger.info(
          `Found total ${allReservations.length} reservations across all chunks`
        );
        await delay(5000);
        try {
          await page.waitForSelector(".tpg-navigation__logo_container", {
            visible: true,
            timeout: 5000,
          });
          await page.click(".tpg-navigation__logo_container");
          await delay(2000); // Wait for navigation
        } catch (error) {
          logger.warn("Could not click navigation logo:", error.message);
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
          "Amount to charge/refund",
          "Reason of charge",
          "Status",
        ],
        ...allReservations.map((res) => [
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
          res.reasonOfCharge || "N/A",
          res.status || "Active",
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
async function processReservationsPage(page, id) {
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

              // Check if this is a canceled reservation
              const isCanceled = await page.evaluate(() => {
                const dialogTitle = document.querySelector(".fds-dialog-title");
                return (
                  dialogTitle &&
                  (dialogTitle.textContent.includes("Cancelled") ||
                    dialogTitle.textContent.includes("Canceled"))
                );
              });

              if (isCanceled) {
                logger.info(
                  "Found canceled reservation, extracting payment information..."
                );
                try {
                  // First, ensure the dialog is visible
                  await page.waitForSelector(".fds-dialog-header", {
                    visible: true,
                    timeout: 5000,
                  });

                  // Wait for payment summary section to be visible
                  await page.waitForSelector(".fds-card-header-title", {
                    visible: true,
                    timeout: 5000,
                  });

                  // Extract payment information from canceled reservation
                  const paymentInfo = await page.evaluate(() => {
                    const getCurrencyValue = (title) => {
                      // Find all section titles
                      const sections = Array.from(
                        document.querySelectorAll(".sidePanelSectionTitle")
                      );
                      const section = sections.find(
                        (el) => el.textContent.trim() === title
                      );
                      if (!section) return "0.00";

                      // Get the currency value from the next cell
                      const valueCell = section
                        .closest(".fds-grid")
                        .querySelector(".fds-currency-value");
                      return valueCell ? valueCell.textContent.trim() : "0.00";
                    };

                    return {
                      cancellationFee: getCurrencyValue("Cancellation fee"),
                      expediaCompensation: getCurrencyValue(
                        "Expedia compensation"
                      ),
                      totalPayout: getCurrencyValue("Your total payout"),
                    };
                  });

                  logger.info(
                    "Extracted payment info for canceled reservation:",
                    paymentInfo
                  );

                  // Add the payment information to the basic data with canceled-specific fields
                  const canceledReservation = {
                    ...basicData,
                    cardNumber: "N/A",
                    expiryDate: "N/A",
                    cvv: "N/A",
                    hasCardInfo: false,
                    hasPaymentInfo: true,
                    totalGuestPayment: "0.00",
                    cancellationFee: paymentInfo.cancellationFee,
                    expediaCompensation: paymentInfo.expediaCompensation,
                    totalPayout: paymentInfo.totalPayout,
                    amountToChargeOrRefund: paymentInfo.cancellationFee,
                    reasonOfCharge: "Cancellation Fee",
                    status: "Cancelled",
                  };

                  pageReservations.push(canceledReservation);
                  logger.info(
                    "Added canceled reservation to results:",
                    canceledReservation
                  );

                  // Try each closing method sequentially with proper waits and checks
                  const closingMethods = [
                    // Method 1: Click the close button using page.click with waitForSelector
                    async () => {
                      const closeButton = await page.waitForSelector(
                        ".fds-dialog-header button.dialog-close",
                        {
                          visible: true,
                          timeout: 2000,
                        }
                      );
                      if (closeButton) {
                        await closeButton.click();
                        return true;
                      }
                      return false;
                    },
                    // Method 2: Use JavaScript click with explicit visibility check
                    async () => {
                      const isSuccess = await page.evaluate(() => {
                        const closeButton = document.querySelector(
                          ".fds-dialog-header button.dialog-close"
                        );
                        if (
                          closeButton &&
                          window.getComputedStyle(closeButton).display !==
                            "none"
                        ) {
                          closeButton.click();
                          return true;
                        }
                        return false;
                      });
                      return isSuccess;
                    },
                    // Method 3: Try alternative close button selector
                    async () => {
                      const altCloseButton = await page.$(
                        '.fds-dialog button[aria-label="Close"]'
                      );
                      if (altCloseButton) {
                        await altCloseButton.click();
                        return true;
                      }
                      return false;
                    },
                    // Method 4: Press Escape key as last resort
                    async () => {
                      await page.keyboard.press("Escape");
                      return true;
                    },
                  ];

                  // Try each method until one succeeds
                  let dialogClosed = false;
                  for (const method of closingMethods) {
                    try {
                      const success = await method();
                      if (success) {
                        // Wait for dialog to be hidden
                        await page
                          .waitForSelector(".fds-dialog-header", {
                            hidden: true,
                            timeout: 3000,
                          })
                          .catch(() => {});

                        // Double check if dialog is really gone
                        const dialogStillVisible = await page.$(
                          ".fds-dialog-header"
                        );
                        if (!dialogStillVisible) {
                          dialogClosed = true;
                          break;
                        }
                      }
                    } catch (methodError) {
                      logger.debug(
                        `Dialog close method failed: ${methodError.message}`
                      );
                      continue;
                    }
                  }

                  if (!dialogClosed) {
                    throw new Error("All dialog closing methods failed");
                  }

                  await delay(1000); // Short stabilization delay
                  continue; // Skip to next reservation
                } catch (error) {
                  logger.warn(
                    `Warning: Could not close canceled reservation dialog: ${error.message}`
                  );
                  continue;
                }
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

              // Get card details with retry mechanism
              let cardData = null;
              let paymentData = null;
              let remainingAmountToCharge = null;
              let amountToRefund = null;
              let retries = 0;
              while (!cardData && !paymentData && retries < 3) {
                try {
                  // First try to get card details
                  cardData = await page.evaluate(() => {
                    const cardNumber =
                      document
                        .querySelector(".cardNumber.replay-conceal bdi")
                        ?.textContent.trim() || "";
                    const expiryDate =
                      document
                        .querySelector(
                          ".cardDetails .fds-cell.all-cell-1-4.fds-type-color-primary.replay-conceal"
                        )
                        ?.textContent.trim() || "";
                    const cvv =
                      document
                        .querySelectorAll(
                          ".cardDetails .fds-cell.all-cell-1-4.fds-type-color-primary.replay-conceal"
                        )[1]
                        ?.textContent.trim() || "";

                    if (cardNumber) {
                      return {
                        cardNumber,
                        expiryDate,
                        cvv,
                      };
                    }
                    return null;
                  });

                  // If no card data, try to get payment information
                  if (!cardData) {
                    paymentData = await page.evaluate(() => {
                      // Find all section titles
                      const sectionTitles = Array.from(
                        document.querySelectorAll(".sidePanelSectionTitle")
                      );

                      // Find the payment sections
                      const totalGuestPaymentTitle = sectionTitles.find((el) =>
                        el.textContent.includes("Total guest payment")
                      );
                      const expediaCompensationTitle = sectionTitles.find(
                        (el) => el.textContent.includes("Expedia compensation")
                      );
                      const totalPayoutTitle = sectionTitles.find((el) =>
                        el.textContent.includes("Your total payout")
                      );

                      // Get the values
                      const totalGuestPayment =
                        totalGuestPaymentTitle?.nextElementSibling
                          ?.querySelector(".fds-currency-value")
                          ?.textContent.trim() || "";
                      const expediaCompensation =
                        expediaCompensationTitle?.nextElementSibling
                          ?.querySelector(".fds-currency-value")
                          ?.textContent.trim() || "";
                      const totalPayout =
                        totalPayoutTitle?.nextElementSibling
                          ?.querySelector(".fds-currency-value")
                          ?.textContent.trim() || "";

                      if (totalGuestPayment) {
                        return {
                          totalGuestPayment,
                          expediaCompensation,
                          totalPayout,
                        };
                      }
                      return null;
                    });
                  }

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

              // Add to reservations array with either card data or payment data
              pageReservations.push({
                ...basicData,
                ...(cardData || {}),
                ...(paymentData || {}),
                hasCardInfo: !!cardData,
                hasPaymentInfo: !!paymentData,
                remainingAmountToCharge: remainingAmountToCharge || "N/A",
                amountToRefund: amountToRefund || "N/A",
                amountToChargeOrRefund:
                  remainingAmountToCharge || amountToRefund || "N/A",
                reasonOfCharge: remainingAmountToCharge
                  ? "Remaining Amount to Charge"
                  : amountToRefund
                  ? "Amount to Refund"
                  : "N/A",
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
                  reasonOfCharge: "N/A",
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
    res.json(data);
  } catch (error) {
    res.status(500).json({ error: "Error reading logs" });
  }
});

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
  if (!loadToken()) {
    logger.info("Opening browser for authentication...");
    open(`http://localhost:${port}/auth`);
  }
});
