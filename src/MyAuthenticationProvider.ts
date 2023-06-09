import {AuthenticationProvider} from "@microsoft/microsoft-graph-client";
import * as fs from "fs";

export class MyAuthenticationProvider implements AuthenticationProvider {
    /**
     * This method will get called before every request to the msgraph server
     * This should return a Promise that resolves to an accessToken (in case of success) or rejects with error (in case of failure)
     * Basically this method will contain the implementation for getting and refreshing accessTokens
     */
    public async getAccessToken(): Promise<string> {
        const token = fs.readFileSync("config/AccessToken", "utf8");
        return token.trim();
    }
}