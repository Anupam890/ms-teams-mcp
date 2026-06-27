import {
  ClientSecretCredential,
  DeviceCodeCredential,
  TokenCredential,
  AccessToken,
  GetTokenOptions,
  AuthenticationRecord,
} from "@azure/identity";
import * as fs from "fs";
import * as path from "path";

export interface AuthConfig {
  tenantId?: string;
  clientId?: string;
  clientSecret?: string;
  authToken?: string;
  authMode: "device-code" | "client-credentials" | "token";
  readOnly?: boolean;
}

const CACHE_DIR = path.join(
  process.env.HOME || process.env.USERPROFILE || ".",
  ".microsoft-teams-mcp"
);
const AUTH_RECORD_FILE = path.join(CACHE_DIR, "auth-record.json");

function loadRecord(): AuthenticationRecord | undefined {
  try {
    if (fs.existsSync(AUTH_RECORD_FILE)) {
      return JSON.parse(fs.readFileSync(AUTH_RECORD_FILE, "utf-8"));
    }
  } catch {}
  return undefined;
}

function saveRecord(record: AuthenticationRecord): void {
  try {
    if (!fs.existsSync(CACHE_DIR)) {
      fs.mkdirSync(CACHE_DIR, { recursive: true });
    }
    fs.writeFileSync(AUTH_RECORD_FILE, JSON.stringify(record, null, 2));
  } catch {}
}

function deleteRecord(): void {
  try {
    if (fs.existsSync(AUTH_RECORD_FILE)) {
      fs.unlinkSync(AUTH_RECORD_FILE);
    }
  } catch {}
}

const DELEGATED_SCOPES = [
  "https://graph.microsoft.com/Team.ReadBasic.All",
  "https://graph.microsoft.com/Channel.ReadBasic.All",
  "https://graph.microsoft.com/ChannelMessage.Send",
  "https://graph.microsoft.com/ChannelMessage.Read.All",
  "https://graph.microsoft.com/Member.Read.All",
  "https://graph.microsoft.com/OnlineMeetings.ReadWrite.All",
  "https://graph.microsoft.com/Calendars.ReadWrite",
  "https://graph.microsoft.com/User.Read.All",
  "https://graph.microsoft.com/Chat.ReadWrite",
  "https://graph.microsoft.com/Chat.Read",
  "https://graph.microsoft.com/Group.Read.All",
  "https://graph.microsoft.com/User.Read",
];
const READ_SCOPES = [
  "https://graph.microsoft.com/User.Read.All",
  "https://graph.microsoft.com/Group.Read.All",
];

export function getScopes(readOnly?: boolean): string[] {
  return readOnly ? READ_SCOPES : DELEGATED_SCOPES;
}

class StaticTokenCredential implements TokenCredential {
  private token: string;
  private expiresOnTimestamp: number;

  constructor(token: string) {
    this.token = token;
    this.expiresOnTimestamp = Infinity;
    try {
      const parts = token.split(".");
      if (parts.length === 3) {
        const payload = JSON.parse(
          Buffer.from(parts[1], "base64").toString("utf-8")
        );
        if (payload.exp) this.expiresOnTimestamp = payload.exp * 1000;
      }
    } catch {}
  }

  async getToken(
    _scopes: string | string[],
    _options?: GetTokenOptions
  ): Promise<AccessToken | null> {
    if (Date.now() >= this.expiresOnTimestamp) {
      throw new Error("AUTH_TOKEN has expired");
    }
    return { token: this.token, expiresOnTimestamp: this.expiresOnTimestamp };
  }
}

export class AuthManager {
  private config: AuthConfig;
  private _credential: TokenCredential | null = null;
  private _deviceCodeCredential: DeviceCodeCredential | null = null;

  constructor(config: AuthConfig) {
    this.config = config;
  }

  private createDeviceCodeCredential(record?: AuthenticationRecord): DeviceCodeCredential {
    return new DeviceCodeCredential({
      tenantId: this.config.tenantId,
      clientId: this.config.clientId,
      authenticationRecord: record,
      tokenCachePersistenceOptions: {
        enabled: true,
        name: "microsoft-teams-mcp",
      },
      userPromptCallback: (info) => {
        console.error("\n" + info.message);
      },
    });
  }

  getCredential(): TokenCredential {
    if (this._credential) return this._credential;

    if (this.config.authToken) {
      this._credential = new StaticTokenCredential(this.config.authToken);
      return this._credential;
    }

    if (this.config.authMode === "client-credentials") {
      if (!this.config.clientSecret || !this.config.tenantId || !this.config.clientId) {
        throw new Error("TEAMS_CLIENT_SECRET, TEAMS_TENANT_ID, and TEAMS_CLIENT_ID are required for client-credentials auth mode");
      }
      this._credential = new ClientSecretCredential(
        this.config.tenantId,
        this.config.clientId,
        this.config.clientSecret
      );
      return this._credential;
    }

    if (!this.config.tenantId || !this.config.clientId) {
      throw new Error("TEAMS_TENANT_ID and TEAMS_CLIENT_ID are required for device-code auth mode");
    }
    const record = loadRecord();
    this._deviceCodeCredential = this.createDeviceCodeCredential(record);
    this._credential = this._deviceCodeCredential;
    return this._credential;
  }

  async isAuthenticated(): Promise<boolean> {
    try {
      const cred = this.getCredential();
      const token = await cred.getToken(getScopes(this.config.readOnly));
      return token !== null && !!token.token;
    } catch {
      return false;
    }
  }

  async ensureAuthenticated(): Promise<void> {
    if (this.config.authMode !== "device-code" || this.config.authToken) {
      const cred = this.getCredential();
      await cred.getToken(getScopes(this.config.readOnly));
      return;
    }
    if (!this.config.tenantId || !this.config.clientId) {
      throw new Error("TEAMS_TENANT_ID and TEAMS_CLIENT_ID are required for device-code auth mode");
    }
    const record = loadRecord();
    if (record) {
      const cred = this.createDeviceCodeCredential(record);
      await cred.getToken(getScopes(this.config.readOnly));
      this._deviceCodeCredential = cred;
      this._credential = cred;
      return;
    }
    const fresh = this.createDeviceCodeCredential();
    const authRecord = await fresh.authenticate(getScopes(this.config.readOnly));
    if (authRecord) saveRecord(authRecord);
    this._deviceCodeCredential = fresh;
    this._credential = fresh;
  }

  async logout(): Promise<void> {
    deleteRecord();
    this._credential = null;
    this._deviceCodeCredential = null;
  }

  getAuthMode(): string {
    if (this.config.authToken) return "token";
    return this.config.authMode;
  }

  isReadOnly(): boolean {
    return this.config.readOnly ?? false;
  }
}
