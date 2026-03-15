// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, expect, it, beforeEach } from "@jest/globals";

// Mock heavy dependencies that auth.ts imports so we can test proxy auth in isolation.
// The jest moduleNameMapper for logger.js conflicts with @azure/logger internals,
// so we must mock the entire dependency chain.
jest.mock("@azure/identity", () => ({}));
jest.mock("@azure/msal-node", () => ({
  PublicClientApplication: jest.fn(),
}));
jest.mock("open", () => jest.fn());
jest.mock("@azure/logger", () => ({
  setLogLevel: jest.fn(),
}));
jest.mock("../../src/logger", () => ({
  logger: {
    debug: jest.fn(),
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
  },
}));

import { createAuthenticator, proxyAuthStore } from "../../src/auth";

describe("proxy authentication", () => {
  describe("createAuthenticator('proxy')", () => {
    let authenticator: () => Promise<string>;

    beforeEach(() => {
      authenticator = createAuthenticator("proxy");
    });

    it("should return a function", () => {
      expect(typeof authenticator).toBe("function");
    });

    it("should return the token from AsyncLocalStorage when available", async () => {
      const result = await proxyAuthStore.run("test-token-123", () => authenticator());
      expect(result).toBe("test-token-123");
    });

    it("should throw when no token is available in AsyncLocalStorage", async () => {
      await expect(authenticator()).rejects.toThrow("No auth token available for this request");
    });

    it("should throw with guidance about _auth_token argument", async () => {
      await expect(authenticator()).rejects.toThrow("_auth_token");
    });

    it("should isolate tokens between concurrent requests", async () => {
      // Simulate two concurrent tool calls with different tokens
      const results = await Promise.all([
        proxyAuthStore.run("user-a-token", async () => {
          // Simulate async work (e.g., API call) before reading token
          await new Promise((resolve) => setTimeout(resolve, 10));
          return authenticator();
        }),
        proxyAuthStore.run("user-b-token", async () => {
          // Simulate async work
          await new Promise((resolve) => setTimeout(resolve, 5));
          return authenticator();
        }),
      ]);

      expect(results[0]).toBe("user-a-token");
      expect(results[1]).toBe("user-b-token");
    });

    it("should support nested async operations within the same context", async () => {
      const result = await proxyAuthStore.run("nested-token", async () => {
        // First async call
        const token1 = await authenticator();
        // Simulate more async work
        await new Promise((resolve) => setTimeout(resolve, 5));
        // Second async call in same context
        const token2 = await authenticator();
        return [token1, token2];
      });

      expect(result).toEqual(["nested-token", "nested-token"]);
    });
  });

  describe("proxyAuthStore", () => {
    it("should return undefined outside of a run context", () => {
      expect(proxyAuthStore.getStore()).toBeUndefined();
    });

    it("should return the stored value inside a run context", () => {
      proxyAuthStore.run("stored-value", () => {
        expect(proxyAuthStore.getStore()).toBe("stored-value");
      });
    });

    it("should not leak state between run contexts", () => {
      proxyAuthStore.run("context-a", () => {
        expect(proxyAuthStore.getStore()).toBe("context-a");
      });

      proxyAuthStore.run("context-b", () => {
        expect(proxyAuthStore.getStore()).toBe("context-b");
      });

      expect(proxyAuthStore.getStore()).toBeUndefined();
    });
  });
});
