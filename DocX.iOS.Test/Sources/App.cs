using System;
using System.IO;

using Foundation;
using UIKit;

namespace DocX.iOS.Test
{
	[Register ("App")]

	public class App : UIApplicationDelegate
	{
		//	Main

		private static void Main (string[] args)
		{
			UIApplication.Main (args, null, "App");
		}

		//	constants

		private const string messageFormat = @"
<html>
<head>
<style type='text/css'>

html, body
{{
	margin: 0;
	padding: 0;
	border: 0;
	font-size: 100%;
	font: inherit;
	font-family: Arial;
	vertical-align: baseline;
	color: #666666;
}}

body
{{
	padding: 2em 2em;
}}

pre
{{
    white-space: pre-wrap;
}}

</style>
</head>
<body>
<h2>{0}</h2>
<pre><code>{1}</code></pre>
</body>
</html>
";
		
		//	publics

		public override UIWindow Window { get; set; }

		//	FinishedLaunching

		public override bool FinishedLaunching (UIApplication application, NSDictionary launchOptions)
		{
			//	create webview + controller

			var controller = new UIViewController ();

			controller.Title = "DocX - Test";

			var webview = new UIWebView (controller.View.Frame);

			webview.AutoresizingMask = UIViewAutoresizing.FlexibleDimensions;
			
			webview.ScalesPageToFit = true;

			controller.View.AddSubview (webview);

			//	create navigation controller

			var navigation = new UINavigationController (controller);

			//	initialize window

			this.Window = new UIWindow (UIScreen.MainScreen.Bounds);

			this.Window.RootViewController = navigation;

			this.Window.MakeKeyAndVisible ();

			//	

			try
			{
				//	path to our temp docx file

				string pathDocx = Path.Combine(Path.GetTempPath (), "Document.docx");

				//	inform user of what we are about to do

				webview.LoadHtmlString (string.Format(messageFormat, "Generating .docx file, please wait...", pathDocx), null);

				//	generating docx

				using (var document = Novacode.DocX.Create (pathDocx))
				{
					Novacode.Paragraph p = document.InsertParagraph();

					p.Append("This is a Word Document");

					p = document.InsertParagraph();

					p.Append("");

					p = document.InsertParagraph();

					p.Append("Hello World");

					document.Save();
				}

				//	showing docx in webview, with delay, otherwise we don't see our initial message

				this.Invoke(() => {

					webview.LoadRequest (NSUrlRequest.FromUrl (NSUrl.FromFilename (pathDocx)));

				}, 2.0f);

				//	done
			}
			catch (Exception e)
			{
				webview.LoadHtmlString (string.Format(messageFormat, "Exception Occurred :", e), null);
			}

			//	done

			return true;
		}
	}
}