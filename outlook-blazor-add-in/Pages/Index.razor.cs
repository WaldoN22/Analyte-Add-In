/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */

using BlazorAddIn.Model;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace BlazorAddIn.Pages
{
	public partial class Index
	{
		[Inject]
		public IJSRuntime JSRuntime { get; set; } = default!;

		public IJSObjectReference JSModule { get; set; } = default!;

		public MailRead? MailReadData { get; set; }

		public CountEmail? CountEmailData { get; set; }

        public UnreadCount? UnEmailData { get; set; }

        /// <summary>
        /// NOTE: This can also go in the @code block in Index.razor
        /// </summary>
        /// <param name = "firstRender" ></ param >
        /// <returns></returns>
        protected override async Task OnAfterRenderAsync(bool firstRender)
		{
			//NOTE: This fires after Index.razor.OnInitialized()
			Console.WriteLine($"Index.razor.cs (OnAfterRenderAsync): firstRender: {firstRender}");

			if (firstRender)
			{
				//NOTE: JSRuntime.InvokeAsync invokes OutlookBlazorWasmApp.Client.lib.module.js(afterStarted), then Index.razor.js(Office.onReady) but only when hosted in full browser instance outside of the Outlook Task pane. When in Outlook, the order after this event completes is:
				//OutlookBlazorWasmApp.Client.lib.module.js(Office.onReady) - after beforeStart has already fired() and triggered Index.razor.OnInitialized()
				//Index.razor.js(Office.onReady) 

				Console.WriteLine($"firstRender: Importing Index.razor.js...");

				JSModule = await JSRuntime.InvokeAsync<IJSObjectReference>("import", "./Pages/Index.razor.js");

				Console.WriteLine($"firstRender: Index.razor.js imported");
			}


			//if (MailReadData == null)
			//{
			//	Console.WriteLine($"========================== Index.razor.cs (OnAfterRenderAsync): Calling GetEmailData()... ==========================");

			//	// Fetch email data from JavaScript function
			//	MailReadData = await GetEmailData();

			//	// Decode Base64 data if available
			//	if (string.IsNullOrEmpty(MailReadData?.AttachmentBase64Data) == false)
			//	{
			//		MailReadData.DecodeBase64();
			//	}

			//	if (MailReadData != null)
			//	{
			//		// Refresh UI with the updated MailReadData
			//		StateHasChanged();
			//	}



			//	Console.WriteLine($"========================== Index.razor.cs (OnAfterRenderAsync): Returning from GetEmailData()... ==========================");
			//}


			if (CountEmailData == null)
			{
				Console.WriteLine($"========================== Index.razor.cs (OnAfterRenderAsync): Calling GetEmailData()... ==========================");

				// Fetch email data from JavaScript function
				CountEmailData = await Count();

				if (CountEmailData != null)
				{
					// Refresh UI with the updated MailReadData
					StateHasChanged();
				}



				Console.WriteLine($"========================== Index.razor.cs (OnAfterRenderAsync): Returning from GetEmailData()... ==========================");
			}

			Console.WriteLine($"Index.razor.cs (OnAfterRenderAsync): Done ...");





            if (UnEmailData == null)
            {
                Console.WriteLine($"========================== Index.razor.cs (OnAfterRenderAsync): Calling GetEmailData()... ==========================");

                // Fetch email data from JavaScript function
                UnEmailData = await UnCount();

                if (UnEmailData != null)
                {
                    // Refresh UI with the updated MailReadData
                    StateHasChanged();
                }



                Console.WriteLine($"========================== Index.razor.cs (OnAfterRenderAsync): Returning from GetEmailData()... ==========================");
            }

            Console.WriteLine($"Index.razor.cs (OnAfterRenderAsync): Done ...");

        }






		//private async Task<MailRead?> GetEmailData()
		//{
		//	// Call JavaScript function to get email data
		//	MailRead? mailreaditem = await JSModule.InvokeAsync<MailRead>("getEmailData");

		//	// Log subject for debugging
		//	Console.WriteLine("Subject C#: ");
		//	Console.WriteLine(mailreaditem?.Subject);

		//	return mailreaditem;
		//}
		private async Task<CountEmail?> Count()
		{
			// Call JavaScript function to get email data
			CountEmail? countMail = await JSModule.InvokeAsync<CountEmail>("countEmailsReceivedToday");

			// Log subject for debugging
			Console.WriteLine("Subject C#: ");
			Console.WriteLine(countMail?.Count);

			return countMail;
		}


        private async Task<UnreadCount?> UnCount()
        {
            // Call JavaScript function to get email data
            UnreadCount? uncountMail = await JSModule.InvokeAsync<UnreadCount>("countUnreadEmails");

            // Log subject for debugging
            Console.WriteLine("Subject C#: ");
            Console.WriteLine(uncountMail?.UnCount);

            return uncountMail;
        }
    }
}
