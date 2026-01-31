const { test, expect } = require('@playwright/test');
const path = require('path');
const xlsx = require('xlsx');

// ================= CONFIGURATION =================
const CONFIG = {
  url: 'https://www.swifttranslator.com/',
  timeouts: {
    pageLoad: 2000,
    afterClear: 1000,
    translation: 3000,
    betweenTests: 2000
  },
  selectors: {
    inputField: 'Input Your Singlish Text Here.',
    outputContainer:
      'div.w-full.h-80.p-3.rounded-lg.ring-1.ring-slate-300.whitespace-pre-wrap'
  }
};

// ================= EXCEL FILE PATH =================
const EXCEL_FILE = path.join(__dirname, '..', 'test_data', 'IT23361386.xlsx');
const SHEET_NAME = 'Test Cases';

// ================= READ EXCEL DATA =================
function readTestData(sheetName) {
  const workbook = xlsx.readFile(EXCEL_FILE);
  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet) {
    throw new Error(`❌ Sheet "${sheetName}" not found in Excel file`);
  }

  const rawData = xlsx.utils.sheet_to_json(worksheet, { defval: '' });

  return rawData.map((row, index) => ({
    rowIndex: index + 2, // +2 because Excel is 1-indexed and has header row
    tcId: row['TC ID'],
    name: row['Test case name'],
    inputType: row['Input length type'],
    input: row['Input'],
    expected: row['Expected output'],
    justification:
      row['Accuracy justification/Description of issue type'],
    coverage: row['What is covered by the test']
  }));
}

// ================= WRITE TO EXCEL =================
function writeTestResult(tcId, actualOutput, status) {
  try {
    const workbook = xlsx.readFile(EXCEL_FILE);
    const worksheet = workbook.Sheets[SHEET_NAME];

    if (!worksheet) {
      console.error(`❌ Sheet "${SHEET_NAME}" not found in Excel file`);
      return;
    }

    // Convert sheet to JSON to find the row
    const rawData = xlsx.utils.sheet_to_json(worksheet, { defval: '', header: 1 });
    
    // Find header row to get column indices
    const headerRow = rawData[0];
    const tcIdColIndex = headerRow.indexOf('TC ID');
    const actualOutputColIndex = headerRow.indexOf('Actual Output');
    const statusColIndex = headerRow.indexOf('Status');

    // If columns don't exist, add them
    let newActualOutputColIndex = actualOutputColIndex;
    let newStatusColIndex = statusColIndex;

    if (actualOutputColIndex === -1) {
      // Add "Actual Output" column after "Expected output"
      const expectedOutputColIndex = headerRow.indexOf('Expected output');
      newActualOutputColIndex = expectedOutputColIndex !== -1 ? expectedOutputColIndex + 1 : headerRow.length;
      headerRow.splice(newActualOutputColIndex, 0, 'Actual Output');
    }

    if (statusColIndex === -1) {
      // Add "Status" column after "Actual Output"
      newStatusColIndex = newActualOutputColIndex + 1;
      headerRow.splice(newStatusColIndex, 0, 'Status');
    }

    // Find the row with matching TC ID
    let targetRowIndex = -1;
    for (let i = 1; i < rawData.length; i++) {
      if (rawData[i][tcIdColIndex] === tcId) {
        targetRowIndex = i;
        break;
      }
    }

    if (targetRowIndex === -1) {
      console.error(`❌ Test case with TC ID "${tcId}" not found in Excel`);
      return;
    }

    // Ensure the row has enough columns
    while (rawData[targetRowIndex].length <= Math.max(newActualOutputColIndex, newStatusColIndex)) {
      rawData[targetRowIndex].push('');
    }

    // Update the row
    rawData[targetRowIndex][newActualOutputColIndex] = actualOutput || '';
    rawData[targetRowIndex][newStatusColIndex] = status || '';

    // Convert back to worksheet
    const newWorksheet = xlsx.utils.aoa_to_sheet(rawData);
    workbook.Sheets[SHEET_NAME] = newWorksheet;

    // Write back to file
    xlsx.writeFile(workbook, EXCEL_FILE);
    console.log(`✅ Updated Excel: TC ${tcId} - Status: ${status}`);
  } catch (error) {
    console.error(`❌ Error writing to Excel for TC ${tcId}:`, error.message);
  }
}

// ================= LOAD TEST DATA =================
const TEST_DATA = readTestData(SHEET_NAME);

console.log(`✅ Loaded ${TEST_DATA.length} test cases from Excel`);

// ================= PAGE OBJECT =================
class TranslatorPage {
  constructor(page) {
    this.page = page;
  }

  async navigateToSite() {
    await this.page.goto(CONFIG.url);
    await this.page.waitForLoadState('networkidle');
    await this.page.waitForTimeout(CONFIG.timeouts.pageLoad);
  }

  async getInputField() {
    return this.page.getByRole('textbox', {
      name: CONFIG.selectors.inputField
    });
  }

  async getOutputField() {
    return this.page
      .locator(CONFIG.selectors.outputContainer)
      .filter({ hasNot: this.page.locator('textarea') })
      .first();
  }

  async clearInput() {
    const input = await this.getInputField();
    await input.fill('');
    await this.page.waitForTimeout(CONFIG.timeouts.afterClear);
  }

  async typeInput(text) {
    const input = await this.getInputField();
    await input.fill(text);
  }

  async waitForOutput() {
    await this.page.waitForFunction(() => {
      const els = Array.from(
        document.querySelectorAll(
          '.w-full.h-80.p-3.rounded-lg.ring-1.ring-slate-300.whitespace-pre-wrap'
        )
      );
      return els.some(el => el.textContent?.trim().length > 0);
    }, { timeout: 10000 });

    await this.page.waitForTimeout(CONFIG.timeouts.translation);
  }

  async getOutputText() {
    const output = await this.getOutputField();
    return (await output.textContent())?.trim();
  }

  async translate(text) {
    await this.clearInput();
    await this.typeInput(text);
    await this.waitForOutput();
    return await this.getOutputText();
  }
}

// ================= TEST SUITE =================
test.describe(
  'SwiftTranslator – Singlish → Sinhala (Excel Driven)',
  () => {
    let translator;

    test.beforeEach(async ({ page }) => {
      translator = new TranslatorPage(page);
      await translator.navigateToSite();
    });

    // Create tests dynamically from Excel data
    TEST_DATA.forEach((tc) => {
      test(`${tc.tcId} - ${tc.name}`, async ({ page }, testInfo) => {
        // Use translator from beforeEach or create new one
        const testTranslator = translator || new TranslatorPage(page);
        if (!translator) {
          await testTranslator.navigateToSite();
        }

        let actualOutput = '';
        let status = 'FAIL';
        let testPassed = false;

        try {
          actualOutput = await testTranslator.translate(tc.input);
          
          // Compare actual with expected
          if (actualOutput === tc.expected) {
            status = 'PASS';
            testPassed = true;
            expect(actualOutput).toBe(tc.expected);
          } else {
            status = 'FAIL';
            testPassed = false;
            // Still write the result even if it fails
            expect(actualOutput).toBe(tc.expected);
          }
        } catch (err) {
          status = 'FAIL';
          testPassed = false;
          
          testInfo.attach('Excel Details', {
            body: `
TC ID       : ${tc.tcId}
Test Name  : ${tc.name}
Input Type : ${tc.inputType}

Input:
${tc.input}

Expected:
${tc.expected}

Actual:
${actualOutput}

Justification:
${tc.justification}

Coverage:
${tc.coverage}
            `,
            contentType: 'text/plain'
          });

          // Don't throw yet - we want to update Excel first
        } finally {
          // Always write results to Excel
          writeTestResult(tc.tcId, actualOutput, status);
          
          if (!testPassed) {
            // Re-throw if test failed so Playwright marks it as failed
            throw new Error(`Test failed: Expected "${tc.expected}" but got "${actualOutput}"`);
          }
        }

        await testTranslator.page.waitForTimeout(
          CONFIG.timeouts.betweenTests
        );
      });
    });
  }
);
