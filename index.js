const puppeteer = require("puppeteer");

const sharepointUrl = "insert URL";
const applicationAddress = "preferrably use the path for msedge.exe";

//example of resourcesObj
const resources = {
  resources: {
    pageExample: {
      component: {
        button: {
          title: "Continue",
        },
      },
    },
  },
};
//example of output:
//path: resources.pageExample.component.button.title
//name: Continue

//this recursive function adds each nested propertyName to it's path.
function getObjectProperties(obj, parentPath = "") {
  let result = [];

  for (const key in obj) {
    if (typeof obj[key] === "object") {
      const currentPath = parentPath ? `${parentPath}.${key}` : key;
      const nestedProperties = getObjectProperties(obj[key], currentPath);
      result = result.concat(nestedProperties);
    } else {
      const currentPath = parentPath ? `${parentPath}.${key}` : key;
      result.push({
        path: currentPath,
        key: key,
        value: obj[key],
      });
    }
  }

  return result;
}

//the 'save' button is separated in a different function because it get's the element by it's text instead of div properties such as type or id
async function clickButtonByText(page, buttonText) {
  await page.evaluate((text) => {
    const buttons = document.querySelectorAll('input[type="button"]');

    for (const button of buttons) {
      if (button.value === text) {
        button.click();
        return;
      }
    }

    console.error(`Button with text '${text}' not found.`);
  }, buttonText);
}

//self explanatory, it fills the forms fields according to conditions.
async function fillContentFields(page, path, name) {
  const maxRetries = 3; // Maximum number of retries

  let retries = 0;

  while (retries < maxRetries) {
    try {
      const contentIdField = '[title="Content ID"]';
      const contentField = '[title="Content"]';
      const titleField = '[title="Title Required Field"]';
      const saveButtonText = "Save";

      // Wait for the 'Content ID' input field to be visible
      await page.waitForSelector(contentIdField);

      await page.type(contentIdField, path);
      await page.waitForTimeout(500);
      await page.type(contentField, name);
      await page.waitForTimeout(500);
      await page.type(titleField, "global");
      await page.waitForTimeout(500);

      try {
        console.log(`Saving`);
        await clickButtonByText(page, saveButtonText);
        console.log("Success!");
      } catch (error) {
        retries++;
        console.error("Error occurred:", error);
        console.log(`Retrying to save... (Attempt ${retries}/${maxRetries})`);
      }

      console.log(`Waiting for changes to take effect for path`);

      await page.waitForTimeout(5000); // Adjust the timeout as needed

      // If everything is successful, break out of the loop
      break;
    } catch (error) {
      console.error("Error occurred:", error);
    }
  }
}

async function run() {
  const browser = await puppeteer.launch({
    headless: false,
    executablePath: applicationAddress,
  });
  const page = await browser.newPage();

  // Navigate to SharePoint and log in (adjust this part based on your authentication method)
  try {
    await page.goto(sharepointUrl);
  } catch (error) {
    //this operation might go wrong by two different ways, but you can manually bypass it by adding your credentials, or follow the message below.
    // would be nice to try to add the error treatment here for clicking in 'Advanced' and typing 'thisisunsafe'
  }

  try {
    const properties = getObjectProperties(resources);

    for (let i = 0; i < properties.length; i++) {
      await page.waitForSelector("#idHomePageNewItem");

      await page.click("#idHomePageNewItem");

      await page.waitForTimeout(5000); // Adjust the timeout as needed
      console.log(
        `property to add: , ${properties[i].path}, ${properties[i].value}`
      );
      await fillContentFields(page, properties[i].path, properties[i].value);
      console.log(`Iteration number ${i + 1} was a success!`);
    }
  } catch (error) {}
  // Close the browser
  //   await browser.close();
}

run();
