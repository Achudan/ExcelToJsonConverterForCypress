# ExcelToJsonConverterForCypress
This repo contains Cypress code and Node js API code, which helps in reading the JSON by converting the EXCEL.

Common problem faced by Automation Testers when they move from Selenium to Cypress is reading test data from Excel. Apache POI provided this feature, So it was easy to play with Excel sheets. To overcome this obstacle in Cypress, this API will be useful.

Note:
1) This API is built to perform READ operation only. It simply converts the Excel into JSON and sent JSON body as reponse.
2) This API supports only XLSX format spreadsheets.

Two small setup is required to perform this operation. One is to run the node application and another is to make API call in Cypress. (DO NOT PANIC. IT IS SIMPLE :) )

# Steps to run node application:
1) Download the files.
2) Make sure node is installed in your machine.
3) Open cmd and navigate to /ExcelToJONAPI_V1.
4) type "node app.js"
5) Now it should be running in 3000 port.

# Steps to integrate in Cypress.
1) Add the below code if you want to read excel data in "BeforeEach" method.

  ```beforeEach(function () {
    cy.request('POST', 'http://localhost:3000/getjson',
    {
      excellocation: "D:\\Projects\\LearnCypress\\cypress\\fixtures\\sample.xlsx",               //excel sheet location
      sheetname: "Sheet1"                                                                        //sheet name
    }).then(function (response) {
      this.data = response.body[0]                                                               //reponse in form of array of objects. response.body[N], where N is row in sheet.
    })
  })
  ```
  2) Add the below code if you want to read excel data anywhere else. (do not forget to add async, await)
  ```
  it("Test Case -Data driven Model ", async function () {
    var postReq = await cy.request('POST', 'http://localhost:3000/getjson',
      {
        excellocation: "D:\\Projects\\LearnCypress\\cypress\\fixtures\\sample.xlsx",
        sheetname: "Sheet1"
      }).then(response => {
        return response.body
      })
    const postResponse = postReq.body
    const firstRowData = postResponse[0]
    cy.visit('https://www.google.com')
    cy.get('.text').type(firstRowData.searchData)
    homePage.selectGender().select(firstRowData.gender)
  })
  ```
