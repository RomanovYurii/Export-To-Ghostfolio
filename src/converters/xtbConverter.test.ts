import fs from "fs";
import ExcelJS from "exceljs";
import { XtbConverter } from "./xtbConverter";
import { SecurityService } from "../securityService";
import { GhostfolioExport } from "../models/ghostfolioExport";
import { GhostfolioOrderType } from "../models/ghostfolioOrderType";
import YahooFinanceServiceMock from "../testing/yahooFinanceServiceMock";

describe("xtbConverter", () => {

  beforeEach(() => {
    jest.spyOn(console, "log").mockImplementation(jest.fn());
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  it("should construct", () => {

    // Act
    const sut = new XtbConverter(new SecurityService(new YahooFinanceServiceMock()));

    // Assert
    expect(sut).toBeTruthy();
  });

  it("should process sample XLSX file", (done) => {

    // Arange
    const sut = new XtbConverter(new SecurityService(new YahooFinanceServiceMock()));
    const inputFile = "samples/xtb-export.xlsx";

    // Act
    sut.readAndProcessFile(inputFile, (actualExport: GhostfolioExport) => {

      // Assert
      expect(actualExport).toBeTruthy();
      expect(actualExport.activities.length).toBeGreaterThan(0);
      expect(actualExport.activities.length).toBe(189);

      // Fee and interest activities must not carry an explicit dataSource:
      // Ghostfolio rejects MANUAL activities with a free-text symbol and
      // resolves the data source for these types itself.
      const nonInvestmentActivities = actualExport.activities
        .filter(activity => activity.type === GhostfolioOrderType.fee || activity.type === GhostfolioOrderType.interest);

      expect(nonInvestmentActivities.filter(activity => activity.type === GhostfolioOrderType.fee).length).toBeGreaterThan(0);
      expect(nonInvestmentActivities.filter(activity => activity.type === GhostfolioOrderType.interest).length).toBeGreaterThan(0);

      nonInvestmentActivities.forEach(activity => {
        expect(activity.dataSource).toBeUndefined();
      });

      done();
    }, (err: Error) => { done(err); });
  });

  it("should process sample CSV file", (done) => {

    // Arange
    const sut = new XtbConverter(new SecurityService(new YahooFinanceServiceMock()));
    const inputFile = "samples/xtb-export.csv";

    // Act
    sut.readAndProcessFile(inputFile, (actualExport: GhostfolioExport) => {

      // Assert
      expect(actualExport).toBeTruthy();
      expect(actualExport.activities.length).toBeGreaterThan(0);
      expect(actualExport.activities.length).toBe(34);

      done();
    }, () => { done.fail("Should not have an error!"); });
  });

  describe("should throw an error if", () => {
    it("the input file does not exist", (done) => {

      // Arrange
      const sut = new XtbConverter(new SecurityService(new YahooFinanceServiceMock()));

      let tempFileName = "tmp/testinput/xtb-filedoesnotexist.csv";

      // Act
      sut.readAndProcessFile(tempFileName, () => { done.fail("Should not succeed!"); }, (err: Error) => {

        // Assert
        expect(err).toBeTruthy();

        done();
      });
    });

    it("the input file is empty", (done) => {

      // Arrange
      const sut = new XtbConverter(new SecurityService(new YahooFinanceServiceMock()));

      let tempFileContent = "";
      tempFileContent += "ID;Type;Time;Symbol;Comment;Amount\n";

      // Act
      sut.processFileContents(tempFileContent, () => { done.fail("Should not succeed!"); }, (err: Error) => {

        // Assert
        expect(err).toBeTruthy();
        expect(err.message).toContain("An error occurred while parsing");

        done();
      });
    });

    it("the header and row column count doesn't match", (done) => {

      // Arrange
      const sut = new XtbConverter(new SecurityService(new YahooFinanceServiceMock()));

      let tempFileContent = "";
      tempFileContent += "ID;Type;Time;Symbol;Comment;Amount\n";

      tempFileContent += `513492358;Stocks/ETF purchase;11.03.2024 10:05:05;SPYL.DE;OPEN BUY 8 @ 11.2835;-90.27;;`;

      // Act
      sut.processFileContents(tempFileContent, () => { done.fail("Should not succeed!"); }, (err: Error) => {

        // Assert
        expect(err).toBeTruthy();
        expect(err.message).toBe("An error occurred while parsing! Details: Invalid Record Length: columns length is 6, got 8 on line 2");

        done();
      });
    });

    it("the XLSX file has no Cash Operations sheet", (done) => {

      // Arrange
      const sut = new XtbConverter(new SecurityService(new YahooFinanceServiceMock()));

      const tempFileName = "tmp/testinput/xtb-nocashoperationssheet.xlsx";

      // Create an XLSX file without a Cash Operations sheet.
      const workbook = new ExcelJS.Workbook();
      workbook.addWorksheet("Open Positions");

      fs.mkdirSync("tmp/testinput", { recursive: true });

      workbook.xlsx.writeFile(tempFileName).then(() => {

        // Act
        sut.readAndProcessFile(tempFileName, () => { done("Should not succeed!"); }, (err: Error) => {

          // Assert
          expect(err).toBeTruthy();
          expect(err.message).toContain("Cash Operations");

          done();
        });
      });
    });

    it("Yahoo Finance throws an error", (done) => {

      // Arrange
      let tempFileContent = "";
      tempFileContent += "ID;Type;Time;Symbol;Comment;Amount\n";
      tempFileContent += `513492358;Stocks/ETF purchase;11.03.2024 10:05:05;SPYL.DE;OPEN BUY 8 @ 11.2835;-90.27`;

      // Mock Yahoo Finance service to throw error.
      const yahooFinanceServiceMock = new YahooFinanceServiceMock();
      jest.spyOn(yahooFinanceServiceMock, "search").mockImplementation(() => { throw new Error("Unit test error"); });
      const sut = new XtbConverter(new SecurityService(yahooFinanceServiceMock));

      // Act
      sut.processFileContents(tempFileContent, () => { done.fail("Should not succeed!"); }, (err: Error) => {

        // Assert
        expect(err).toBeTruthy();
        expect(err.message).toContain("Unit test error");

        done();
      });
    });
  });

  it("should log when Yahoo Finance returns no symbol", (done) => {

    // Arrange
    let tempFileContent = "";
    tempFileContent += "ID;Type;Time;Symbol;Comment;Amount\n";
    tempFileContent += `513492358;Stocks/ETF purchase;11.03.2024 10:05:05;SPYL.DE;OPEN BUY 8 @ 11.2835;-90.27`;

    // Mock Yahoo Finance service to return no quotes.
    const yahooFinanceServiceMock = new YahooFinanceServiceMock();
    jest.spyOn(yahooFinanceServiceMock, "search").mockImplementation(() => { return Promise.resolve({ quotes: [] }) });
    const sut = new XtbConverter(new SecurityService(yahooFinanceServiceMock));

    // Bit hacky, but it works.
    const consoleSpy = jest.spyOn((sut as any).progress, "log");

    // Act
    sut.processFileContents(tempFileContent, () => {

      expect(consoleSpy).toHaveBeenCalledWith("[i] No result found for action buy, symbol SPYL.DE and comment OPEN BUY 8 @ 11.2835! Please add this manually..\n");

      done();
    }, () => done.fail("Should not have an error!"));
  });

  it("should log error and invoke errorCallback when an error occurs in processFileContents", (done) => {

    // Arrange
    const tempFileContent = "ID;Type;Time;Symbol;Comment;Amount\n";
    const sut = new XtbConverter(new SecurityService(new YahooFinanceServiceMock()));

    const consoleSpy = jest.spyOn(console, "log");

    // Act
    sut.processFileContents(tempFileContent, () => {
      done.fail("Should not succeed!");
    }, (err: Error) => {

      // Assert
      expect(consoleSpy).toHaveBeenCalledWith("[e] An error occurred while processing the file contents. Stack trace:");
      expect(consoleSpy).toHaveBeenCalledWith(err.stack);
      expect(err).toBeTruthy();

      done();
    });
  });
});
