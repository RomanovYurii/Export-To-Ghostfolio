import fs from "fs";
import GhostfolioService from "./ghostfolioService";

describe("ghostfolioService", () => {

  const inputFile = "tmp/testinput/ghostfolio-export.json";

  let fetchSpy: jest.SpyInstance;

  function mockResponse(status: number, body: any): Response {
    return {
      ok: status >= 200 && status < 300,
      status,
      json: () => Promise.resolve(body)
    } as unknown as Response;
  }

  beforeEach(() => {
    jest.spyOn(console, "log").mockImplementation(jest.fn());

    process.env.GHOSTFOLIO_URL = "http://localhost:3333";
    process.env.GHOSTFOLIO_SECRET = "unit-test-secret";
    process.env.GHOSTFOLIO_VALIDATE = "true";
    process.env.GHOSTFOLIO_IMPORT = "true";

    fs.mkdirSync("tmp/testinput", { recursive: true });
    fs.writeFileSync(inputFile, JSON.stringify({ activities: [{ type: "BUY" }] }));

    fetchSpy = jest.spyOn(globalThis, "fetch");
  });

  afterEach(() => {
    fetchSpy.mockRestore();
    jest.clearAllMocks();

    delete process.env.GHOSTFOLIO_URL;
    delete process.env.GHOSTFOLIO_SECRET;
    delete process.env.GHOSTFOLIO_VALIDATE;
    delete process.env.GHOSTFOLIO_IMPORT;
  });

  it("should authenticate with a POST request containing the access token", (done) => {

    // Arrange
    fetchSpy.mockImplementation((input: any, init?: any) => {
      if (`${input}`.includes("/auth/anonymous")) {
        return Promise.resolve(mockResponse(200, { authToken: "unit-test-token" }));
      }

      return Promise.resolve(mockResponse(201, { activities: [] }));
    });

    const sut = new GhostfolioService();

    // Act
    sut.validate(inputFile).then((result: boolean) => {

      // Assert
      expect(result).toBe(true);
      expect(fetchSpy).toHaveBeenCalledTimes(2);

      const [authUrl, authInit] = fetchSpy.mock.calls[0];
      expect(`${authUrl}`).toBe("http://localhost:3333/api/v1/auth/anonymous");
      expect(authInit.method).toBe("POST");
      expect(JSON.parse(authInit.body)).toEqual({ accessToken: "unit-test-secret" });

      const [importUrl, importInit] = fetchSpy.mock.calls[1];
      expect(`${importUrl}`).toBe("http://localhost:3333/api/v1/import?dryRun=true");
      expect(importInit.headers).toEqual(expect.arrayContaining([["Authorization", "Bearer unit-test-token"]]));

      done();
    }).catch((err: Error) => { done(err); });
  });

  it("should stop retrying when authentication keeps failing", (done) => {

    // Arrange
    fetchSpy.mockImplementation((input: any) => {
      if (`${input}`.includes("/auth/anonymous")) {
        return Promise.resolve(mockResponse(200, { authToken: "unit-test-token" }));
      }

      return Promise.resolve(mockResponse(401, { message: "Unauthorized" }));
    });

    const sut = new GhostfolioService();

    // Act
    sut.validate(inputFile).then(() => {

      done(new Error("Should not succeed!"));
    }).catch((err: Error) => {

      // Assert
      expect(err.message).toContain("authentication error");
      expect(fetchSpy.mock.calls.length).toBeLessThanOrEqual(7);

      done();
    });
  });

  it("should fail fast when the authentication request does not return a token", (done) => {

    // Arrange
    fetchSpy.mockImplementation(() => Promise.resolve(mockResponse(404, { message: "Cannot GET" })));

    const sut = new GhostfolioService();

    // Act
    sut.validate(inputFile).then(() => {

      done(new Error("Should not succeed!"));
    }).catch((err: Error) => {

      // Assert
      expect(err.message).toContain("Failed to authenticate with Ghostfolio");
      expect(err.message).toContain("(HTTP 404)");
      expect(fetchSpy).toHaveBeenCalledTimes(1);

      done();
    });
  });

  it("should return the activity count after a successful import", (done) => {

    // Arrange
    fetchSpy.mockImplementation((input: any) => {
      if (`${input}`.includes("/auth/anonymous")) {
        return Promise.resolve(mockResponse(200, { authToken: "unit-test-token" }));
      }

      return Promise.resolve(mockResponse(201, { activities: [{}, {}] }));
    });

    const sut = new GhostfolioService();

    // Act
    sut.import(inputFile).then((count: number) => {

      // Assert
      expect(count).toBe(2);

      done();
    }).catch((err: Error) => { done(err); });
  });

  it("should throw after a failed import and log a single message", (done) => {

    // Arrange
    fetchSpy.mockImplementation((input: any) => {
      if (`${input}`.includes("/auth/anonymous")) {
        return Promise.resolve(mockResponse(200, { authToken: "unit-test-token" }));
      }

      return Promise.resolve(mockResponse(400, { message: "activities.0.symbol is not valid" }));
    });

    const consoleSpy = jest.spyOn(console, "log");
    const sut = new GhostfolioService();

    // Act
    sut.import(inputFile).then(() => {

      done(new Error("Should not succeed!"));
    }).catch((err: Error) => {

      // Assert
      expect(err.message).toBe("Automatic import failed! See the logs for more details.");
      expect(consoleSpy).toHaveBeenCalledWith("[e]\tactivities.0.symbol is not valid");

      done();
    });
  });
});
