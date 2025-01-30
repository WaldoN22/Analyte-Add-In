export async function beforeStart(wasmoptions, extensions) {
    console.log("beforeStart: BlazorAddIn.lib.module.js()");
    console.log("beforeStart: Entering function");

    try {
        // Ensure Office.js is ready before proceeding
        await Office.onReady();
        console.log("beforeStart: Office.js is ready.");

        // Check if we're running inside Outlook
        if (Office.context.mailbox) {
            console.log("beforeStart: Running inside Outlook Task Pane.");
            console.log("Mailbox context:", Office.context.mailbox);
        } else {
            console.log("beforeStart: Running inside a web browser.");
        }
    } catch (error) {
        console.error("beforeStart: Error initializing Office.js", error);
    }
}

/**
 * Called after Blazor is ready to receive calls from JS.
 * @param  {} blazor
 */
export async function afterStarted(blazor) {
    console.log("afterStarted: BlazorAddIn.lib.module.js():");
    console.log("afterStarted: Entering function");
    try {
        // Check if Blazor is initialized correctly
        console.log("Blazor:", blazor);
    } catch (error) {
        console.error("afterStarted: Error with Blazor initialization", error);
    }
}
