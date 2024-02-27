const { Builder, Browser, By, Key, until } = require("selenium-webdriver");
const xl = require("exceljs");
(async function example() {
  let driver = await new Builder().forBrowser(Browser.FIREFOX).build();
  try {
    let meta_data = {
      service_index: 7,
      service_name: "مصلحة الرجال",
      en_name: "men_surgery",
      start_date: "12/1/2023",
      end_date: "1/1/2024",
      month: "december",
    };
    await driver.get("http://172.16.1.2:800/Calender/Index");

    // const data = [
    //   {
    //     code: "M0339",
    //     name: "ساشي خيط صفراء صغيرة",
    //     qty: "1",
    //     unitCost: "86",
    //   },
    //   {
    //     code: "M0267",
    //     name: "صابون سائل غسول اليدين",
    //     qty: "1",
    //     unitCost: "69.9",
    //   },
    //   { code: "M0268", name: "معطر الجو", qty: "1", unitCost: "190" },
    //   { code: "M0264", name: "ممسحة ارضيات", qty: "2", unitCost: "71" },
    //   { code: "M0224", name: "مناديل مبللة", qty: "1", unitCost: "80" },
    //   {
    //     code: "M0149",
    //     name: "مناديل ورقيه (بابي سيرفات)",
    //     qty: "30",
    //     unitCost: "32",
    //   },
    //   {
    //     code: "5154",
    //     name: "ساشي بوبيل سوداء غردائية",
    //     qty: "1",
    //     unitCost: "220",
    //   },
    //   {
    //     code: "5153",
    //     name: "ساشي بوبيل صفراء حجم كبير",
    //     qty: "1",
    //     unitCost: "169.95",
    //   },
    //   {
    //     code: "M0276",
    //     name: "ساشي سوداء خيط حجم صغير",
    //     qty: "1",
    //     unitCost: "86",
    //   },
    //   { code: "M0268", name: "معطر الجو", qty: "1", unitCost: "190" },
    //   { code: "M0114", name: "سائل الاواني", qty: "1.5", unitCost: "60" },
    //   { code: "M111", name: "صانيبوا", qty: "3.5", unitCost: "30" },
    //   { code: "M0419", name: "RAM 90*90", qty: "2", unitCost: "45" },
    //   { code: "M0433", name: "RAM A3 160G", qty: "4", unitCost: "1800" },
    //   { code: "M0396", name: "RAM A4 ENTET", qty: "2", unitCost: "2365" },
    //   { code: "M0329", name: "قلم ارزق", qty: "5", unitCost: "22.5" },
    //   { code: "M0252", name: "DVD VIERGE", qty: "150", unitCost: "57" },
    //   { code: "M0391", name: "وسادة", qty: "2", unitCost: "480" },
    //   {
    //     code: "5154",
    //     name: "ساشي بوبيل سوداء غردائية",
    //     qty: "15",
    //     unitCost: "0",
    //   },
    //   {
    //     code: "M0378",
    //     name: "ساشي خيط صفراء حجم كبير",
    //     qty: "1",
    //     unitCost: "0",
    //   },
    //   {
    //     code: "M0339",
    //     name: "ساشي خيط صفراء صغيرة",
    //     qty: "15",
    //     unitCost: "0",
    //   },
    //   {
    //     code: "M0276",
    //     name: "ساشي سوداء خيط حجم صغير",
    //     qty: "15",
    //     unitCost: "0",
    //   },
    //   {
    //     code: "M0267",
    //     name: "صابون سائل غسول اليدين",
    //     qty: "1",
    //     unitCost: "0",
    //   },
    //   { code: "M0297", name: "ليقو ميناج", qty: "1", unitCost: "0" },
    //   { code: "M0268", name: "معطر الجو", qty: "2", unitCost: "0" },
    //   { code: "M0264", name: "ممسحة ارضيات", qty: "2", unitCost: "0" },
    //   {
    //     code: "100032",
    //     name: "ENVELOPE F30 CLILIQUE IBN HAYANE",
    //     qty: "500",
    //     unitCost: "14.86",
    //   },
    //   {
    //     code: "M0309",
    //     name: "ENVELOPE RADIO",
    //     qty: "25",
    //     unitCost: "67.5",
    //   },
    //   { code: "M0433", name: "RAM A3 160G", qty: "4", unitCost: "1800" },
    //   {
    //     code: "M0359",
    //     name: "بوشات ايكوغرافي",
    //     qty: "1",
    //     unitCost: "38.9",
    //   },
    //   {
    //     code: "M0149",
    //     name: "مناديل ورقيه (بابي سيرفات)",
    //     qty: "29",
    //     unitCost: "32",
    //   },
    //   { code: "M0252", name: "DVD VIERGE", qty: "3", unitCost: "0" },
    //   {
    //     code: "M0399",
    //     name: "TONER (NOIR/BLANCHE)",
    //     qty: "1",
    //     unitCost: "0",
    //   },
    //   { code: "M0398", name: "رلولو كود بار", qty: "2", unitCost: "0" },
    //   {
    //     code: "M0378",
    //     name: "ساشي خيط صفراء حجم كبير",
    //     qty: "1",
    //     unitCost: "0",
    //   },
    //   {
    //     code: "M0339",
    //     name: "ساشي خيط صفراء صغيرة",
    //     qty: "1",
    //     unitCost: "0",
    //   },
    //   {
    //     code: "M0276",
    //     name: "ساشي سوداء خيط حجم صغير",
    //     qty: "1",
    //     unitCost: "0",
    //   },
    //   {
    //     code: "M0346",
    //     name: "شريط لاصق حجم صغير",
    //     qty: "1",
    //     unitCost: "0",
    //   },
    //   {
    //     code: "M0303",
    //     name: "صابون حجرة سائل 3L",
    //     qty: "1",
    //     unitCost: "0",
    //   },
    //   {
    //     code: "M0267",
    //     name: "صابون سائل غسول اليدين",
    //     qty: "2",
    //     unitCost: "0",
    //   },
    //   { code: "M0405", name: "عمارة جرافيز", qty: "2", unitCost: "0" },
    //   { code: "M0329", name: "قلم ارزق", qty: "6", unitCost: "0" },
    //   { code: "M0342", name: "قلم اسود", qty: "2", unitCost: "0" },
    //   { code: "M0323", name: "كأس بلاستيك", qty: "96", unitCost: "0" },
    //   { code: "M0268", name: "معطر الجو", qty: "2", unitCost: "0" },
    //   { code: "M0264", name: "ممسحة ارضيات", qty: "1", unitCost: "0" },
    //   { code: "M109", name: "ممسحة الغبار", qty: "1", unitCost: "0" },
    //   { code: "M0224", name: "مناديل مبللة", qty: "2", unitCost: "0" },
    //   { code: "M0309", name: "ENVELOPE RADIO", qty: "1400", unitCost: "0" },
    //   { code: "M0252", name: "DVD VIERGE", qty: "2", unitCost: "2750" },
    //   { code: "M0114", name: "سائل الاواني", qty: "1", unitCost: "0" },
    //   {
    //     code: "M0267",
    //     name: "صابون سائل غسول اليدين",
    //     qty: "1",
    //     unitCost: "0",
    //   },
    //   { code: "M0268", name: "معطر الجو", qty: "2", unitCost: "0" },
    //   { code: "M0114", name: "سائل الاواني", qty: "1", unitCost: "0" },
    //   {
    //     code: "M0267",
    //     name: "صابون سائل غسول اليدين",
    //     qty: "1",
    //     unitCost: "0",
    //   },
    //   { code: "M0268", name: "معطر الجو", qty: "2", unitCost: "0" },
    // ];

    // await createXlTable(data);

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
      .then(async (_) => await driver.sleep(5000))
      .then(
        async (_) =>
          await stimListItemClick(
            "#saved-filters > div > div > div:nth-child(2) > div > span.k-widget.k-combobox.k-combobox-clearable > span > span.k-select > span",
            `#search-filters_listbox > li:nth-child(${meta_data.service_index})`
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
          meta_data.start_date
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
          meta_data.end_date
        );
      })
      .then(async (_) => await driver.sleep(3000))
      .then(async (_) => {
        return;
        await driver.executeScript(
          `arguments[0].click()`,
          await driver.findElement(
            By.css(
              'a[title="Go to the last page"][aria-label="Go to the last page"]'
            )
          )
        );
      })
      .then(async (_) => {
        const indicators = await driver.findElements(
          By.css(
            '#MasterGrid a[aria-label="Go to the previous page"][title="Go to the previous page"] + div > ul > li'
          )
        );
        return indicators;
      })
      .then(async (indicators) => {
        const items = [];
        for (const [i, ind] of indicators.entries()) {
          ind.click();
          await driver.sleep(2000);
          const rowCount = await getElements(
            `#saved-filters + div > #MasterGrid > div:nth-child(2) > table > tbody > tr`
          );
          for (const [j, row] of rowCount.entries()) {
            await fetchData(row).then((data) => items.push(...data));
          }
        }
        return items;
      })
      .then(
        async (items) =>
          await createXlTable(items, meta_data.en_name, meta_data.month)
      );

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
        `arguments[0].dispatchEvent(new KeyboardEvent('keypress', {
                key: '${key}',
                keyCode: ${keyCode},
                bubbles: true,
                cancelable: true
            }));`,
        await driver.findElement(By.css(ele))
      );
    }

    async function fetchData(row) {
      return new Promise(async (resolve, reject) => {
        const data = [];

        try {
          await driver.executeScript(
            `arguments[0].click()`,
            await row.findElement(By.css(`td:first-child > a:first-child`))
          );

          await driver.sleep(6000);

          const rowCount = await getElements(
            "#dvTranferReq + div.row div.k-grid-content > table > tbody > tr"
          );

          for (const row of rowCount) {
            let obj = {
              code: "",
              name: "",
              qty: 0,
              unitCost: 0,
            };

            const name_code = await row
              .findElement(By.css("td:nth-child(7)"))
              .getAttribute("innerHTML");
            const [name, code] = name_code.split("|");
            obj.code = code.trim();
            obj.name = name.trim();

            const qty = await row
              .findElement(By.css("td:nth-child(9)"))
              .getAttribute("innerHTML");
            obj.qty = qty.trim();

            const unitCost = await row
              .findElement(By.css("td:nth-child(10)"))
              .getAttribute("innerHTML");
            obj.unitCost = unitCost.trim();
            data.push(obj);
          }

          await driver.sleep(6000);

          await driver.executeScript(
            `arguments[0].click()`,
            await driver.findElement(By.css("a.close"))
          );

          await driver.sleep(2000);

          resolve(data); // Resolve the Promise with the data
        } catch (error) {
          reject(error); // Reject the Promise if there's an error
        }
      });
    }

    async function getElements(ele) {
      const rowCount = await driver.findElements(By.css(ele));
      return rowCount;
    }

    async function createXlTable(data, service, month) {
      const wb = new xl.Workbook();
      const ws = wb.addWorksheet("sheet1");
      ws.columns = [
        { header: "Code", key: "code", width: 10 },
        { header: "Name", key: "name", width: 20 },
        { header: "Qty", key: "qty", width: 20 },
        { header: "UnitCost", key: "unitCost", width: 20 },
      ];
      data.forEach((item) => ws.addRow(item));
      const filePath = `C:\\Users\\WIN10_CIH\\Documents\\Excel\\transaction_items_${service}_${month}.xlsx`;
      wb.xlsx
        .writeFile(filePath)
        .then((_) => console.log(`Excel file has been created successfully`));
    }
  } catch (e) {
    console.error(e);
  } finally {
    // await driver.quit();
  }
})();
