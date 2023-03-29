import { Middleware } from "@microsoft/microsoft-graph-client";
import { Context } from "@microsoft/microsoft-graph-client";

export class MyLoggingHandler implements Middleware {
    private nextMiddleware: Middleware;

    public async execute(context: Context): Promise<void> {
        try {
            let url: string;
            if (typeof context.request === "string") {
                url = context.request;
            } else {
                url = context.request.url;
            }
            console.log(url);
            return await this.nextMiddleware.execute(context);
        } catch (error) {
            throw error;
        }
    }

    public setNext(next: Middleware): void {
        this.nextMiddleware = next;
    }
}