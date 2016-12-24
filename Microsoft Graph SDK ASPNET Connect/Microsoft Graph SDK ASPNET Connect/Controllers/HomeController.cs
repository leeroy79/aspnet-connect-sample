/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using System.Threading.Tasks;
using System.Web.Mvc;
using Microsoft.Graph;
using Microsoft_Graph_SDK_ASPNET_Connect.Helpers;
using Microsoft_Graph_SDK_ASPNET_Connect.Models;
using Resources;
using System;
using System.Text;

namespace Microsoft_Graph_SDK_ASPNET_Connect.Controllers
{
    public class HomeController : Controller
    {
        GraphService graphService = new GraphService();

        public ActionResult Index()
        {
            return View("Graph");
        }

        [Authorize]
        // Get the current user's email address from their profile.
        public async Task<ActionResult> GetMyEmailAddress()
        {
            try
            {

                // Get an access token.
                string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync();

                // Get the current user's email address. 
                ViewBag.Email = await graphService.GetMyEmailAddress(accessToken);
                return View("Graph");
            }
            catch (Exception e)
            {
                if (e.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = Resource.Error_Message + Request.RawUrl + ": " + e.Message });
            }
        }

        [Authorize]
        // Send mail on behalf of the current user.
        public async Task<ActionResult> SendEmail()
        {
            if (string.IsNullOrEmpty(Request.Form["email-address"]))
            {
                ViewBag.Message = Resource.Graph_SendMail_Message_GetEmailFirst;
                return View("Graph");
            }

            // Build the email message.
            Microsoft_Graph_SDK_ASPNET_Connect.Models.MessageRequest email = graphService.BuildEmailMessage(Request.Form["recipients"], Request.Form["subject"]);

            try
            {

                // Get an access token.
                string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync();

                // Send the email.
                ViewBag.Message = await graphService.SendEmail(accessToken, email);

                // Reset the current user's email address and the status to display when the page reloads.
                ViewBag.Email = Request.Form["email-address"];
                return View("Graph");
            }
            catch (Exception e)
            {
                if (e.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = Resource.Error_Message + Request.RawUrl + ": " + e.Message });
            }
        }



        [Authorize]
        // Send mail on behalf of the current user.
        public async Task<ActionResult> GetMyTasks()
        {
            try
            {
                string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync();

                var myTasks = await graphService.GetMyTasks(accessToken);

                var sb = new StringBuilder();

                foreach(var task in myTasks)
                {
                    sb.Append(task.Title);
                    sb.AppendLine("<br/>");
                }

                ViewBag.Message = sb.ToString();

                return View("Graph");
            }
            catch (Exception e)
            {
                if (e.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new
                {
                    message = Resource.Error_Message + Request.RawUrl + ": " + e.Message
                });
            }
        }




        public ActionResult About()
        {
            return View();
        }
    }
}