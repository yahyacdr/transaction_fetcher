const { Builder, Browser, By, Key, until } = require("selenium-webdriver");
const notifier = require("node-notifier");
const { StaleElementReferenceError } = require("selenium-webdriver").error;
const xl = require("exceljs");
(async function example() {
  let driver = await new Builder().forBrowser(Browser.FIREFOX).build();
  try {
    let meta_data = {
      service_index: [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18],
      service_name: ["المطبخ", "مصلحة التوليد"],
      en_name: "men_surgery",
      start_date: "12/31/2023",
      end_date: "02/01/2024",
      month: "January",
    };
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
    }
    async function switchToNewTab() {
      await driver.switchTo().newWindow("tab");
      await driver.get("http://172.16.1.2:800/InventoryTransferRequest/index");
    }
    await logIn();
    for (let [i, s_i] of meta_data.service_index.entries()) {
      await switchToNewTab();
      await fetch_transaction_data(s_i);
      notifier.notify({
        title: "Final Task Done",
        message: "Fetch job finished Successfully",
        // Add an icon (optional)
        icon: "path/to/icon.png", // Optional, absolute path to an icon
        // Add a sound (optional)
        sound: true, // Only Notification Center or Windows Toasters
        wait: true, // Wait with callback, until user action is taken against notification
      });
    }
    // return;
    async function fetch_transaction_data(s_i) {
      await driver
        .sleep(500)
        .then(async (_) => {
          await stimListItemClick(
            "#saved-filters > div > div > div:nth-child(2) > div > span.k-widget.k-combobox.k-combobox-clearable > span > span.k-select > span",
            `#search-filters_listbox > li:nth-child(${s_i})`
          );
          meta_data.service_name = await driver
            .findElement(
              By.css(`#search-filters_listbox > li:nth-child(${s_i})`)
            )
            .getText();
        })
        .then(async (_) => {
          await clickItem(
            await driver.findElement(
              By.css(
                "#MasterGrid > div:first-child table > thead > tr:first-child> th:nth-child(6) > a > span"
              )
            ),
            "#MasterGrid > div:first-child table > thead > tr:first-child> th:nth-child(6) > a > span"
          );
        })
        .then(async (_) => {
          await stimEvent(
            "body > div:last-child ul > li.k-item.k-menu-item.k-filter-item.k-state-default.k-last",
            "mouseover"
          );
        })
        .then(async (_) => {
          await clickItem(
            await driver.findElement(
              By.css(
                "body > div:last-child ul > li.k-item.k-menu-item.k-filter-item.k-state-default.k-last form .k-filter-menu-container span.k-widget.k-dropdown"
              )
            ),
            "body > div:last-child ul > li.k-item.k-menu-item.k-filter-item.k-state-default.k-last form .k-filter-menu-container span.k-widget.k-dropdown"
          );
        })
        // .then(async (_) => await driver.sleep(800))
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
          await clickItem(
            await driver.findElement(
              By.css('span[title="Additional operator"]')
            ),
            'span[title="Additional operator"]'
          );
        })
        .then(async (_) => {
          await clickItem(
            await driver.findElement(
              By.css(
                "body > div:last-child ul > li.k-item.k-menu-item.k-filter-item.k-state-default.k-last > .k-animation-container > ul > div.k-animation-container:nth-child(5) ul > li:nth-child(6)"
              )
            ),
            "body > div:last-child ul > li.k-item.k-menu-item.k-filter-item.k-state-default.k-last > .k-animation-container > ul > div.k-animation-container:nth-child(5) ul > li:nth-child(6)"
          );
        })
        // .then(async (_) => await driver.sleep(500))
        .then(async (_) => {
          await stimInputType(
            "body > div:last-child ul > li.k-item.k-menu-item.k-filter-item.k-state-default.k-last > .k-animation-container > ul > li:first-child form > div > span:nth-child(6) input",
            meta_data.end_date
          );
        })
        // .then(async (_) => await driver.sleep(2000))
        .then(
          async (_) =>
            await driver
              .findElement(
                By.css(
                  `button[type='submit'][title='Filter'].k-button.k-primary`
                )
              )
              .click()
        )
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
            const li = `#MasterGrid a[aria-label="Go to the previous page"][title="Go to the previous page"] + div > ul > ${
              i === 0 ? `li:first-child` : `li:nth-child(${i + 1}) a`
            }`;
            await clickItem(await driver.findElement(By.css(li)), li);
            // await driver.sleep(3000);
            const rowCount = await getElements(
              `div#saved-filters + div > div#MasterGrid > div:nth-child(2) > table > tbody > tr`
            );
            for (const [j, row] of rowCount.entries()) {
              await fetchData(
                row,
                `form#InvRequestFrm > div#dvTranferReq + div div#Grid > div:nth-child(3) > table > tbody > tr`,
                j + 1
              ).then((data) => items.push(...data));
            }
          }
          return items;
        })
        .then(async (items) => {
          // const mergedItems = [];
          // items.forEach((item) => {
          //   // Check if the code already exists in the mergedItems array
          //   const existingItem = mergedItems.find(
          //     (element) => element.code === item.code
          //   );

          //   if (!existingItem) {
          //     mergedItems.push({ ...item });
          //   } else {
          //     existingItem.qty = (
          //       parseInt(existingItem.qty) + parseInt(item.qty)
          //     ).toString();
          //     existingItem.unitCost = (
          //       parseInt(existingItem.unitCost) + parseInt(item.unitCost)
          //     ).toString();
          //   }
          // });
          await createTemplate(
            items,
            meta_data.start_date,
            meta_data.end_date,
            meta_data.service_name
          );
        })
        .then(async () => {
          notifier.notify({
            title: "Task Done",
            message: "Fetch job finished Successfully",
            // Add an icon (optional)
            icon: "path/to/icon.png", // Optional, absolute path to an icon
            // Add a sound (optional)
            sound: true, // Only Notification Center or Windows Toasters
            wait: true, // Wait with callback, until user action is taken against notification
          });
        });
    }

    async function stimInputTypeWList(ele, value, list_ele) {
      await clickItem(await driver.findElement(By.css(ele)), ele);
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
      await clickItem(await driver.findElement(By.css(list_ele)), list_ele);
    }

    async function stimListItemClick(listBtn, item) {
      await clickItem(await driver.findElement(By.css(listBtn)), listBtn);
      await clickItem(await driver.findElement(By.css(item)), item);
    }

    async function stimInputType(ele, value) {
      await clickItem(await driver.findElement(By.css(ele)), ele);
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

    async function fetchData(row, css, j) {
      return new Promise(async (resolve, reject) => {
        const data = [];

        try {
          await clickItem(
            await row.findElement(By.css(`td:first-child > a:first-child`)),
            `div#saved-filters + div > div#MasterGrid > div:nth-child(2) > table > tbody > tr:nth-child(${j})`
          );

          // await driver.sleep(7000);

          const rowCount = await getElements(css);

          for (const [r, row] of rowCount.entries()) {
            let obj = {
              code: "",
              name: "",
              qty: 0,
              unitCost: 0,
            };

            await waitForElement(
              `${css}:nth-child(${r + 1}) > td:nth-child(7)`
            );

            const name_code = await driver
              .findElement(
                By.css(`${css}:nth-child(${r + 1}) > td:nth-child(7)`)
              )
              .getAttribute("innerHTML");
            const [name, code] = name_code.split("|");
            obj.code = code.trim();
            obj.name = name.trim();

            waitForElement(`${css}:nth-child(${r + 1}) > td:nth-child(9)`);

            const qty = await driver
              .findElement(
                By.css(`${css}:nth-child(${r + 1}) > td:nth-child(9)`)
              )
              .getAttribute("innerHTML");
            obj.qty = qty.trim();

            await waitForElement(
              `${css}:nth-child(${r + 1}) > td:nth-child(10)`
            );

            const unitCost = await driver
              .findElement(
                By.css(`${css}:nth-child(${r + 1}) > td:nth-child(10)`)
              )
              .getAttribute("innerHTML");
            obj.unitCost = unitCost.trim();
            data.push(obj);
          }

          // await driver.sleep(3000);

          await clickItem(
            await driver.findElement(By.css("a.close")),
            "a.close"
          );

          // await driver.sleep(2000);

          resolve(data); // Resolve the Promise with the data
        } catch (error) {
          reject(error); // Reject the Promise if there's an error
        }
      });
    }

    async function getElements(ele) {
      let elements;
      for (let retry = 0; retry < 6; retry++) {
        // Retry up to 3 times
        try {
          // Locate the elements
          await driver.wait(until.elementsLocated(By.css(ele)), 60000);
          elements = await driver.findElements(By.css(ele));

          // Wait for one of the elements to become visible
          await driver.wait(until.elementIsVisible(elements[0]), 60000);

          // Return the elements if they are still valid
          await driver.findElements(By.css(ele)); // Ensure elements are still valid
          return elements;
        } catch (error) {
          if (error instanceof StaleElementReferenceError) {
            console.log("Stale element reference encountered, retrying...");
          } else {
            throw error; // Re-throw other errors
          }
        }
      }
      throw new Error("Failed to locate elements after multiple retries");
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

    async function createTemplate(items, start_date, end_date, service_name) {
      await driver.executeScript(
        (items, start_date, end_date, service_name) => {
          let report = `<head>
          <meta charset="UTF-8" />
          <meta name="viewport" content="width=device-width, initial-scale=1.0" />
          <title>Report</title>
        </head>
        
        <body>
          <style>
            /*
            * Prefixed by https://autoprefixer.github.io
            * PostCSS: v8.4.14,
            * Autoprefixer: v10.4.7
            * Browsers: last 4 version
            */
        
            @import url("https://fonts.googleapis.com/css2?family=Changa:wght@200;300;400;500;600;700;800&display=swap");
        
            *,
            *::before,
            *::after {
              -webkit-box-sizing: border-box;
              box-sizing: border-box;
              padding: 0;
              margin: 0;
              font-family: "Changa", sans-serif;
            }
        
            html {
              font-size: 12px;
            }
        
            .report {
              width: 1080.06px;
              width: 720.36px;
              /* height: 507.626px; */
              background-color: window;
              padding: 0;
              position: relative;
        
              input {
                border: none;
                text-align: center;
              }
              input[type="number"] {
                width: 80px;
                font-size: 1rem;
              }
        
              header {
                background-color: black;
                display: -webkit-box;
                display: -ms-flexbox;
                display: flex;
                -webkit-box-pack: center;
                -ms-flex-pack: center;
                justify-content: center;
                -webkit-box-align: center;
                -ms-flex-align: center;
                align-items: center;
                width: 100%;
                height: 12vh;
                padding: 10px 10px;
        
                .title-date {
                  display: -webkit-box;
                  display: -ms-flexbox;
                  display: flex;
                  -webkit-box-orient: vertical;
                  -webkit-box-direction: normal;
                  -ms-flex-direction: column;
                  flex-direction: column;
                  -webkit-box-pack: justify;
                  -ms-flex-pack: justify;
                  justify-content: space-between;
                  -webkit-box-align: center;
                  -ms-flex-align: center;
                  align-items: center;
                  color: white;
                  width: 45%;
                  direction: rtl;
                  height: 100%;
        
                  .date {
                    width: 75%;
                    display: -webkit-box;
                    display: -ms-flexbox;
                    display: flex;
                    -webkit-box-pack: justify;
                    -ms-flex-pack: justify;
                    justify-content: space-between;
                    -webkit-box-align: center;
                    -ms-flex-align: center;
                    align-items: center;
                    padding-top: 30px;
                    font-size: 1.05rem;
                  }
                }
              }
        
              main {
                display: -webkit-box;
                display: -ms-flexbox;
                display: flex;
                -webkit-box-orient: vertical;
                -webkit-box-direction: normal;
                -ms-flex-direction: column;
                flex-direction: column;
                -webkit-box-pack: justify;
                -ms-flex-pack: justify;
                justify-content: space-between;
                -webkit-box-align: center;
                -ms-flex-align: center;
                align-items: center;
                position: relative;
        
                .top {
                  width: 100%;
                  padding: 30px 0;
                  display: -webkit-box;
                  display: -ms-flexbox;
                  display: flex;
                  -webkit-box-pack: center;
                  -ms-flex-pack: center;
                  justify-content: center;
                  -webkit-box-align: center;
                  -ms-flex-align: center;
                  align-items: center;
        
                  > div {
                    direction: rtl;
                    width: 100%;
                    display: -webkit-box;
                    display: -ms-flexbox;
                    display: flex;
                    -webkit-box-pack: center;
                    -ms-flex-pack: center;
                    justify-content: center;
                    -webkit-box-align: center;
                    -ms-flex-align: center;
                    align-items: center;
                    font-size: 1.4rem;
                    font-weight: 600;
                  }
                  input {
                    text-align: center;
                    font-size: 1.4rem;
                    font-weight: 600;
                  }
                }
        
                table {
                  direction: rtl;
                  border: 1px solid black;
                  width: 100%;
                  border-spacing: 0;
        
                  thead {
                    th:not(:last-child) {
                      border-left: 1px solid black;
                    }
        
                    th {
                      padding: 10px;
                      background-color: #eeece1;
                      font-size: 1rem;
                    }
        
                    th:first-child {
                      width: 11%;
                    }
                    th:nth-child(2) {
                      width: 33%;
                    }
                    th:nth-child(3) {
                      width: 16%;
                    }
                    th:nth-child(4) {
                      width: 16%;
                    }
                    th:nth-child(5) {
                      width: 24%;
                    }
                  }
        
                  & tbody {
                    tr {
                      td {
                        text-align: center;
                        padding: 10px;
                        font-size: 1rem;
                      }
        
                      td:first-child {
                        input {
                          width: 50px;
                        }
                      }
        
                      td:not(:last-child) {
                        border-left: 1px solid black;
                      }
        
                      textarea {
                        overflow-wrap: break-word;
                        width: 100%;
                        height: 27px;
                        text-align: justify;
                        resize: none;
                        border: none;
                        overflow-y: hidden;
                      }
                    }
        
                    tr:first-child {
                      td {
                        border-top: 1px solid black;
                      }
                    }
        
                    tr:not(:last-child) {
                      td {
                        border-bottom: 1px solid black;
                      }
                    }
                    .highlight {
                      -webkit-box-shadow: 0 0 10px 3px rgba(0, 0, 0, 0.2);
                      box-shadow: 0 0 10px 3px rgba(0, 0, 0, 0.2);
                    }
                  }
                  & tfoot {
                    display: table-footer-group;
                    vertical-align: middle;
                    border: inherit;
                    tr {
                      td {
                        text-align: center;
                        padding: 10px;
                        font-size: 1rem;
                        border-top: 1px solid black;
                      }
                      td:not(:last-child) {
                        border-left: 1px solid black;
                      }
                    }
                  }
                }
                .add-row-btn {
                  position: absolute;
                  bottom: 0;
                  right: 0;
                  width: 15px;
                  height: 15px;
                  background-color: transparent;
                  button {
                    width: 100%;
                    height: 100%;
                    border: none;
                    background-color: transparent;
                    position: relative;
                    top: 0px;
                    cursor: crosshair;
                  }
                }
              }
        
              .context-menu {
                list-style: none;
                width: 300px;
                display: none;
                -webkit-box-pack: start;
                -ms-flex-pack: start;
                justify-content: flex-start;
                -webkit-box-align: center;
                -ms-flex-align: center;
                align-items: center;
                position: absolute;
                background-color: #171717;
                border-radius: 4px;
                padding: 10px 10px;
                li {
                  color: white;
                  font-size: 1rem;
                  width: 100%;
                  cursor: pointer;
                  padding: 4px;
                  display: -webkit-box;
                  display: -ms-flexbox;
                  display: flex;
                  -webkit-box-pack: justify;
                  -ms-flex-pack: justify;
                  justify-content: space-between;
                  -webkit-box-align: center;
                  -ms-flex-align: center;
                  align-items: center;
                  border-radius: 4px;
                  img {
                    width: 18px;
                    height: 18px;
                  }
                }
                li:hover {
                  background-color: #212121;
                }
              }
        
              footer {
                display: -webkit-box;
                display: -ms-flexbox;
                display: flex;
                -webkit-box-pack: center;
                -ms-flex-pack: center;
                justify-content: center;
                -webkit-box-align: end;
                -ms-flex-align: end;
                align-items: flex-end;
                padding: 70px 0 0;
        
                > div {
                  display: -webkit-box;
                  display: -ms-flexbox;
                  display: flex;
                  -webkit-box-pack: justify;
                  -ms-flex-pack: justify;
                  justify-content: space-between;
                  -webkit-box-align: center;
                  -ms-flex-align: center;
                  align-items: center;
                  width: 50%;
        
                  p,
                  input {
                    font-size: 1.2rem;
                    font-weight: bold;
                  }
                }
              }
            }
          </style>
        
          <div class="report">
            <header>
              <div class="title-date">
                <h1>التقرير الربع سنوي للإستهلاك</h1>
                <div class="date">
                  <p>من:&nbsp;</p>
                  <p class="theDate">${start_date}</p>
                  <p>إلى:&nbsp;</p>
                  <p class="theDate">${end_date}</p>
                </div>
              </div>
            </header>
            <main>
              <div class="top">
                <div class="to-wrh">
                  <p>قسم:&nbsp;</p>
                  <input type="text" class="toWrh" value="الأشعة" />
                </div>
              </div>
              <table>
                <thead>
                  <tr>
                    <th>رمز الصنف</th>
                    <th>اسم الصنف</th>
                    <th>الكمية</th>
                    <th>قيمة الوحدة</th>
                    <th>القيمة</th>
                  </tr>
                </thead>
                <tbody>
                  <tr data-i="1">
                    <td class="code"><input type="text" value="M0296" /></td>
                    <td class="name">
                      <input type="text" value="TERMOMETRE INFLA" />
                    </td>
                    <td class="qty">
                      <input type="number" value="01" />
                    </td>
                    <td class="price-unit"><input type="number" /></td>
                    <td class="price"><input type="number" /></td>
                  </tr>
                </tbody>
                <tfoot>
                  <tr>
                    <td colspan="2">المجموع</td>
                    <td class="qty-sum">55</td>
                    <td class="price-unit-sum">55</td>
                    <td class="price-sum">55</td>
                  </tr>
                </tfoot>
              </table>
              <div class="add-row-btn">
                <button type="button"></button>
              </div>
              <ol class="context-menu">
                <li>
                  Delete
                  <img src="https://i.ibb.co/CK4VCj2/bin.png" alt="bin" border="0" />
                </li>
              </ol>
            </main>
            <footer>
              <div class="sign">
                <p>المخزني</p>
                <p><input type="text" value="المستلم" /></p>
              </div>
            </footer>
          </div>
        </body>
        `;
          let html = document.querySelector("html");
          html.className = "";
          while (html.firstChild) {
            html.removeChild(html.firstChild);
          }
          html.innerHTML = report;
          let body = document.querySelector("body");
          let addRowBtn = body.querySelector("main .add-row-btn button");
          let contextmenu = body.querySelector(".context-menu");
          let deleteBtn = body.querySelector("ol li");

          document.querySelector("input.toWrh").value = service_name;
          let rowInstance = document
            .querySelector("table > tbody > tr:first-child")
            .cloneNode(true);
          let tbody = document.querySelector("table > tbody");

          addRowBtn.addEventListener("click", addRow);

          function addRow() {
            tbody.appendChild(rowInstance);
            rowInstance = tbody.querySelector("tr").cloneNode(true);
          }

          let qtySum = 0;
          let priceSum = 0;
          items.forEach((p) => {
            rowInstance.querySelector(".code input").value = p.code;
            rowInstance.querySelector(".name input").value = p.name;
            rowInstance.querySelector(".qty input").value =
              p.qty < 10 ? "0" + p.qty : p.qty;
            rowInstance.querySelector(".price-unit input").value =
              p.unitCost < 10 ? "0" + p.unitCost : p.unitCost;
            rowInstance.querySelector(".price input").value =
              p.unitCost * p.qty < 10
                ? "0" + p.unitCost * p.qty
                : p.unitCost * p.qty;
            tbody.appendChild(rowInstance);
            rowInstance = tbody.querySelector("tr").cloneNode(true);
            qtySum += p.qty;
            priceSum += p.unitCost * p.qty;
          });

          arrangeRows();

          document.querySelector("tfoot .qty-sum").innerHTML = qtySum;
          document.querySelector("tfoot .price-sum").innerHTML = priceSum;

          document.addEventListener("keydown", function (event) {
            if (event.altKey && event.key === "+") {
              addRow();
            }
          });

          let allInputs = body.querySelectorAll("table input");

          allInputs.forEach((input) => {
            input.addEventListener("input", (e) => {
              if (e.target.parentElement.classList.contains("qty")) {
                document.querySelector("tfoot .qty-sum").innerHTML = countSum(
                  ".qty > input",
                  false
                );
              }
              if (e.target.parentElement.classList.contains("price")) {
                document.querySelector("tfoot .price-sum").innerHTML = countSum(
                  ".price > input",
                  false
                );
              }
              if (e.target.parentElement.classList.contains("price-unit")) {
                e.target.parentElement.parentElement.querySelector(
                  ".price input"
                ).value =
                  +e.target.parentElement.parentElement.querySelector(
                    ".qty input"
                  ).value * +e.target.value;
                document.querySelector("tfoot .price-sum").innerHTML = countSum(
                  ".price > input",
                  false
                );
              }
            });
          });

          function countSum(className, both) {
            if (!both) {
              const sum = Array.from(
                document.querySelectorAll(className),
                (input) => +input.value
              ).reduce((curr, acc) => {
                return curr + acc;
              }, 0);
              return sum;
            } else {
              document.querySelector(".qty-sum").innerHTML = Array.from(
                document.querySelectorAll(".qty > input"),
                (input) => +input.value
              )
                .reduce((curr, acc) => {
                  return curr + acc;
                }, 0)
                .toFixed(2);

              document.querySelector(".price-sum").innerHTML = Array.from(
                document.querySelectorAll(".price > input"),
                (input) => +input.value
              )
                .reduce((curr, acc) => {
                  return curr + acc;
                }, 0)
                .toFixed(2);
            }
          }

          tbody.addEventListener("contextmenu", (e) => {
            let rows = tbody.querySelectorAll("tr");
            if (e.target.tagName === "TD") {
              e.preventDefault();
              rows.forEach((row) => row.classList.remove("highlight"));
              e.target.parentElement.classList.add("highlight");
              contextmenu.style = `display: flex; top: ${e.clientY}px; left: ${e.clientX}px`;
              contextmenu.setAttribute(
                "current-i",
                e.target.parentElement.getAttribute("data-i")
              );
            }
          });

          document.addEventListener("click", function (event) {
            const specificElement = tbody.querySelectorAll(".highlight");

            specificElement.forEach((se) => {
              if (!se.contains(event.target) && event.target !== se) {
                se.classList.remove("highlight");
                contextmenu.style.display = "none";
              }
            });
          });

          deleteBtn.addEventListener("click", () => {
            let deleteRow = tbody.querySelector(
              `tr[data-i="${deleteBtn.parentElement.getAttribute(
                "current-i"
              )}"]`
            );
            deleteRow.parentElement.removeChild(deleteRow);
            contextmenu.style.display = "none";
            arrangeRows();
            countSum("", true);
          });

          function arrangeRows() {
            let rows = tbody.querySelectorAll("tbody tr");
            rows.forEach((row, i) => {
              row.setAttribute("data-i", i + 1 < 10 ? "0" + (i + 1) : i + 1);
            });
          }

          tbody.querySelector("tr").remove();
        },
        items,
        start_date,
        end_date,
        service_name
      );
    }

    async function scrollItemToView(item) {
      await driver.executeScript("arguments[0].scrollIntoView(true);", item);
      await driver.sleep(500);
    }

    async function clickItem(item, css) {
      await waitForElement(css);
      await scrollItemToView(item);
      await driver.executeScript(`arguments[0].click()`, item);
    }

    async function waitForElement(css) {
      await driver.wait(until.elementsLocated(By.css(css)), 60000);

      await driver.wait(
        until.elementIsVisible(await driver.findElement(By.css(css))),
        60000
      );
    }
  } catch (e) {
    console.error(e);
    notifier.notify({
      title: "Task Failed",
      message: "Error occurred while fetching",
      // Add an icon (optional)
      icon: "path/to/icon.png", // Optional, absolute path to an icon
      // Add a sound (optional)
      sound: true, // Only Notification Center or Windows Toasters
      wait: true, // Wait with callback, until user action is taken against notification
    });
  } finally {
    // await driver.quit();
  }
})();
