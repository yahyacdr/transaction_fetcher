const { Builder, Browser, By, Key, until } = require("selenium-webdriver");

(async function example() {
  let driver = await new Builder().forBrowser(Browser.FIREFOX).build();
  try {
    await driver.get("http://172.16.1.2:800/Calender/Index");

    // ENTER USERNAME & PASSWORD
    async function logIn() {
      await driver
        .findElement(By.id("username"))
        .sendKeys("magazine", Key.RETURN);
      await driver.findElement(By.id("password")).sendKeys("9865", Key.RETURN);

      // SUBMIT
      await driver.findElement(By.id("Submit1")).click();

      await driver.sleep(2000);

      await driver.get("http://172.16.1.2:800/InventoryTransferRequest/index");
    }
    await logIn()
      .then(async (_) => await driver.sleep(3000))
      .then(
        async (_) =>
          await stimListItemClick(
            "#saved-filters > div > div > div:nth-child(2) > div > span.k-widget.k-combobox.k-combobox-clearable > span > span.k-select > span",
            "#search-filters_listbox > li:last-child"
          )
      )
      .then(async (_) => {
        await driver.executeScript(
          `arguments[0].click()`,
          await driver.findElement(
            By.css(
              "#MasterGrid > div:first-child table > thead > tr:first-child> th:nth-child(6) > a > span"
            )
          )
        );
      })
      .then(async (_) => {
        await stimEvent(
          "body > div:last-child ul > li.k-item.k-menu-item.k-filter-item.k-state-default.k-last",
          "mouseover"
        );
      })
      .then(async (_) => {
        await driver.executeScript(
          `arguments[0].click()`,
          await driver.findElement(
            By.css(
              "body > div:last-child ul > li.k-item.k-menu-item.k-filter-item.k-state-default.k-last form .k-filter-menu-container span.k-widget.k-dropdown"
            )
          )
        );
      })
      .then(async (_) => await driver.sleep(800))
      .then(async (_) => {
        await stimEvent(
          "body > div:last-child ul > li.k-item.k-menu-item.k-filter-item.k-state-default.k-last > .k-animation-container ul > li:nth-child(4)",
          "click"
        );
      })
      .then(async (_) => {
        await stimInputType(
          "body > div:last-child ul > li.k-item.k-menu-item.k-filter-item.k-state-default.k-last > .k-animation-container form > div > span:nth-child(3) input",
          "12/1/2023"
        );
      })
      .then(async (_) => {
        await driver.executeScript(
          `arguments[0].click()`,
          await driver.findElement(By.css('span[title="Additional operator"]'))
        );
      })
      .then(async (_) => {
        await driver.executeScript(
          `arguments[0].click()`,
          await driver.findElement(
            By.css(
              "body > div:last-child ul > li.k-item.k-menu-item.k-filter-item.k-state-default.k-last > .k-animation-container > ul > div.k-animation-container:nth-child(5) ul > li:nth-child(6)"
            )
          )
        );
      })
      .then(async (_) => {
        stimInputType(
          "body > div:last-child ul > li.k-item.k-menu-item.k-filter-item.k-state-default.k-last > .k-animation-container > ul > li:first-child form > div > span:nth-child(6) input",
          "1/1/2024"
        );
      })
      .then(async (_) => {
        stimKeyPress(
          "body > div:last-child ul > li.k-item.k-menu-item.k-filter-item.k-state-default.k-last > .k-animation-container > ul > li:first-child form > div > span:nth-child(6) input",
          "Enter",
          13
        );
      });

    async function stimInputTypeWList(ele, value, list_ele) {
      await driver.executeScript(
        `arguments[0].click()`,
        await driver.findElement(By.css(ele))
      );
      await driver.sleep(1000);
      await driver.executeScript(
        `arguments[0].value = '${value}'`,
        await driver.findElement(By.css(ele))
      );
      await driver.executeScript(
        `arguments[0].dispatchEvent(new Event('input', { bubbles: true }))`,
        await driver.findElement(By.css(ele))
      );
      await driver.sleep(1000);
      await driver.executeScript(
        `arguments[0].click()`,
        await driver.findElement(By.css(list_ele))
      );
    }

    async function stimListItemClick(listBtn, item) {
      await driver
        .executeScript(
          `arguments[0].click()`,
          await driver.findElement(By.css(listBtn))
        )
        .then(async (_) => {
          await driver.executeScript(
            `arguments[0].click()`,
            await driver.findElement(By.css(item))
          );
        });
    }

    async function stimInputType(ele, value) {
      await driver.executeScript(
        `arguments[0].click()`,
        await driver.findElement(By.css(ele))
      );
      await driver.sleep(500);
      await driver.executeScript(
        `arguments[0].value = '${value}'`,
        await driver.findElement(By.css(ele))
      );
      await driver.sleep(200);
      await driver.executeScript(
        `arguments[0].dispatchEvent(new Event('change', { bubbles: true }))`,
        await driver.findElement(By.css(ele))
      );
    }

    async function stimEvent(ele, eve) {
      await driver.executeScript(
        `arguments[0].dispatchEvent(new MouseEvent('${eve}', { bubbles: true, cancelable: true, view: window }));`,
        await driver.findElement(By.css(ele))
      );
    }

    async function stimKeyPress(ele, key, keyCode) {
      await driver.executeScript(
        `arguments[0].dispatchEvent(new KeyboardEvent('keydown', {
                key: '${key}',
                keyCode: ${keyCode},
                bubbles: true,
                cancelable: true
            }));`,
        await driver.findElement(By.css(ele))
      );
    }
  } catch (e) {
    console.error(e);
  } finally {
    // await driver.quit();
  }
})();
