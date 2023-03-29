import {Context, Middleware} from "@microsoft/microsoft-graph-client";

export class MyHttpMessageHandler implements Middleware {
    public async execute(context: Context): Promise<void> {
        try {
            // For more information about context object refer "Context" section below
            let response = await fetch(context.request, context.options);
            // Set the response back in the context
            context.response = response;
            return;
        } catch (error) {
            throw error;
        }
    }
}