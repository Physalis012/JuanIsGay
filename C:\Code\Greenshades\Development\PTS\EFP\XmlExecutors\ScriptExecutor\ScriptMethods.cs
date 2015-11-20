using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Reflection;
using System.Xml;
using System.Threading;
using System.IO;
using Greenshades.EFP;
using System.Windows.Forms;
using System.Diagnostics;
using System.Windows.Automation;
using Microsoft.Win32;
using Greenshades.EFP.EFPSubmitter;
using Greenshades.BrowserDriver;

namespace Greenshades.EFP {

    public partial class ScriptExecutor {
        
        #region Available Commands

        [Method("Adds a Get Value that will NOT store for the specified type.  Ex if the submission is a Filing and DontStoreForType is selected, any Get() with the specified value will not be stored in the submission parameters used for submit.")]
        public void AddDontStore(string value, FilingPayment DontStoreForType) {
            Log();
            // this method does all its work in the submission parameter collector, so we dont actually need to do anything here
            // its just a stub to tell the collector what to do when it runs on this script file
            return;
        }

        [Method()]
        public void SetStateForTotalColumns(string value) {
            Log();
            // this is just a place holder for automated totals in the SubmissionParameterCollector
        }

        [Method("Sets a default value for a Submission Parameter.  This is important if you are going to later update the submission parameter but it didnt exist in the original Submission Parameters.")]
        public void SetSubmissionParameterDefault(string name, VariableType variableType, string value) {
            Log();
            object retVal = Get(name);

            if (retVal == null) {
                UpdateSubmissionParameter(name, variableType, value);
            }
        }

        [Method("Clicks the correct TAP withholding account links")]
        public void SelectTAPWithholdingPayAccount(string accountNumber, string tableID, FilingFrequency filingFrequency = FilingFrequency.Quarterly) {

            try {
                Browser.FindElement(BrowserDriver.ElementType.Table, ElementAttributes.ID, tableID);
                Browser.FindInTable(SearchType.Contains, BrowserDriver.ElementType.TableRow, ElementAttributes.InnerText, accountNumber,
                    SearchType.Contains, ElementAttributes.InnerText, "Withholding");

                Browser.FindElementInElement(BrowserDriver.ElementType.Anchor, ElementAttributes.Any);
                if (Browser.Element != null) {
                    Browser.Element.Click();
                }
                else {
                    OnError("The submitter could not find the correct withholding account specified.", null, ErrorCategory.Data, ErrorAssignee.Support);
                }              

                if (filingFrequency == FilingFrequency.Annual) {
                    return;
                }
            }
            catch (Exception ex) {
                OnError("Error selecting correct TAP eServices account.", ex);
            }
        }


        [Method("Clicks the correct withholding period's pay links for TAP system sites")]
        public void SelectTAPWithholdingPayPeriod(DateTime reportingPeriodEndDate, string tableID, FileOrPay fileOrPay, string frequency) {
            try {
                DateTime periodEnd = reportingPeriodEndDate;

                FilingFrequency filingFrequency = FilingFrequency.Quarterly;
                try {
                    filingFrequency = (FilingFrequency)Enum.Parse(typeof(FilingFrequency), frequency);
                }
                catch {
                    //if it doesn't parse as a valid frequency Quarterly is the default
                }
                
                if (filingFrequency == FilingFrequency.Quarterly) {
                    periodEnd = GetRelativeDate(reportingPeriodEndDate, RelativeDateType.LastDayOfQuarter);
                }
                else if (filingFrequency == FilingFrequency.Monthly) {
                    periodEnd = GetRelativeDate(reportingPeriodEndDate, RelativeDateType.LastDayOfMonth);
                }
                else if (filingFrequency == FilingFrequency.Annual) {
                    periodEnd = GetRelativeDate(reportingPeriodEndDate, RelativeDateType.LastDayOfYear);
                }

                string textToSelect = "";

                if (fileOrPay == FileOrPay.Pay) {
                    textToSelect = "Pay";
                }
                else if (fileOrPay == FileOrPay.File) {
                    textToSelect = "File Now";
                }

                try {
                    Browser.FindElement(BrowserDriver.ElementType.Table, ElementAttributes.ID, tableID);
                    Browser.FindInTable(SearchType.Contains, BrowserDriver.ElementType.TableRow, ElementAttributes.InnerText, textToSelect,
                       SearchType.Contains, ElementAttributes.InnerText, periodEnd.ToString("dd-MMM-yyy"));

                    if (!Browser.Element.IsNull) { //the date wasn't the correct format so try the second format
                        Browser.FindElement(BrowserDriver.ElementType.Table, ElementAttributes.ID, tableID);
                        Browser.FindInTable(SearchType.Contains, BrowserDriver.ElementType.TableRow, ElementAttributes.InnerText, textToSelect,
                           SearchType.Contains, ElementAttributes.InnerText, periodEnd.ToString("MMM-dd-yyy"));
                    }
                }
                catch (Exception ex) {
                    OnError("Could not find element.  " + ex.Message, ex);
                    return;
                }

                Browser.Element.Child(BrowserDriver.ElementType.TableCell).Child(BrowserDriver.ElementType.Div).Child(BrowserDriver.ElementType.Label).Child(BrowserDriver.ElementType.Anchor).Click();

            }
            catch (Exception ex) {
                OnError("Error selecting correct TAP eServices payment period.", ex);
            }
        }

        [Method("Updates a Submission parameter when the script finishes running in preparation for a later run.")]
        public void UpdateSubmissionParameter(string name, VariableType variableType, string value) {
            Log();
            UpdateSubmissionParameter subParam = new EFP.UpdateSubmissionParameter();
            subParam.Name = name;
            subParam.Value = value;


            switch (variableType) {
                case VariableType.Bool:
                    subParam.Type = typeof(bool);
                    break;
                case VariableType.DateTime:
                    subParam.Type = typeof(DateTime);
                    break;
                case VariableType.Decimal:
                    subParam.Type = typeof(decimal);
                    break;
                case VariableType.Integer:
                    subParam.Type = typeof(int);
                    break;
                case VariableType.String:
                    subParam.Type = typeof(string);
                    break;
            }
            subParam.EncryptionKeyType = null; 
            
            if (ScriptResult.UpdateSubmissionParameters.Contains(subParam)) {
                ScriptResult.UpdateSubmissionParameters.Remove(subParam);
            }
            ScriptResult.UpdateSubmissionParameters.Add(subParam);
            SetVariable(name, variableType, value);
        }


        [Method("Confirms that the agency is still processing the filing or payment in the middle of the Submission")]
        public void ConfirmProcessingSubmission(int waitMinutes = 0, int waitHours = 0, int waitDays = 0) {
            Log();
            try {
                ScriptResult.ScriptStatus = EFPSubmitter.SubmissionStatus.ProcessingSubmission;
                DateTime nextTryDate = DateTime.UtcNow;
                nextTryDate = nextTryDate.AddDays(waitDays);
                nextTryDate = nextTryDate.AddHours(waitHours);
                nextTryDate = nextTryDate.AddMinutes(waitMinutes);
                ScriptResult.NextTryUTC = nextTryDate;
                SaveScreenshot();
            }
            catch (Exception ex) {
                OnError("Error setting submitter result status to 'ProcessingSubmission exectuing the script", ex, ErrorCategory.Script, ErrorAssignee.Support, 8042);
            }
        }

        [Method("Confirms that a payment of filing got submitted successfully.  Searches for the 'html' specified in the first parameter, and then saves the follow  x characters specified by the second parameter as the agencyComfirmationID")]
        public void ConfirmSubmitted(string htmlDirectlyBeforeTextToSave, int numCharactersToSave) {
            Log();
            try {
                string originalSearchText = htmlDirectlyBeforeTextToSave;
                string htmlToSearch = Browser.Html;

                // sometimes this thing doesn't give me the page's html for some reason, im trying to track this error down
                if (htmlToSearch == "<BODY></BODY>") {
                    htmlToSearch = Browser.Html;
                }

                Match match = Regex.Match(htmlToSearch, Regex.Escape(htmlDirectlyBeforeTextToSave) + ".*", RegexOptions.IgnoreCase);
                if (!match.Success) {
                    htmlDirectlyBeforeTextToSave = htmlDirectlyBeforeTextToSave.Replace("\"", "");
                    match = Regex.Match(htmlToSearch, htmlDirectlyBeforeTextToSave + ".*", RegexOptions.IgnoreCase);
                }
                if (match.Success) {
                    string eFileID = match.Value.Substring(htmlDirectlyBeforeTextToSave.Length, Math.Min(numCharactersToSave, match.Value.Length - htmlDirectlyBeforeTextToSave.Length)).Trim();

                    ScriptResult.AgencyConfirmationID = eFileID;
                    ScriptResult.ScriptStatus = EFPSubmitter.SubmissionStatus.Submitted;
                    SaveScreenshot();
                    if(ScriptResult.Screenshot != null) {
                        ScriptResult.ReceiptFile = new SubmissionFile("Receipt_" + eFileID + ".png", "Submission Receipt", "Receipt", "Image", ScriptResult.Screenshot);                        
                    }
                }
                else {
                    string message = "While trying to parse the confirmationID, I could not find any html/text matching " + originalSearchText + " in the page html, so I tried replacing the quotes and got " + htmlDirectlyBeforeTextToSave +
                        " but still couldn't find a match.";
                    OnError(message, null, ErrorCategory.Script, ErrorAssignee.Support, 8017);
                }
            }
            catch (Exception ex) {
                OnError("Errored while trying to parse the confirmation ID from the webpage after submitting.", ex, ErrorCategory.Script, ErrorAssignee.Support, 8017);
            }
        }

        [Method("Confirms that a payment or filing got submitted successfully.  Saves the agency ConfirmationID set in the 'literalAgencyConfirmationID' argument.  Only use this method when doing tests.")]
        public void ConfirmSubmitted2(string literalAgencyConfirmationID) {
            Log();
            try {

                if(string.IsNullOrEmpty(literalAgencyConfirmationID)){
                    try {
                        string html = Browser.Html;                        
                        Trace.WriteLine(DateTime.Now.ToString());
                        Trace.Indent();
                        Trace.WriteLine(html);
                        Trace.Unindent();                       
                    }
                    catch {
                    }
                    OnError("The confirmationID was null or empty when calling Confirm Submitted2.  See log information for page html.");
                    return;
                }

                ScriptResult.AgencyConfirmationID = literalAgencyConfirmationID;
                ScriptResult.ScriptStatus = EFPSubmitter.SubmissionStatus.Submitted;
                SaveScreenshot();
                Log("screenshot saved");
                if (ScriptResult.Screenshot != null) {
                    ScriptResult.ReceiptFile = new SubmissionFile("Receipt_" + literalAgencyConfirmationID + ".png", "Submission Receipt", "Receipt", "Image", ScriptResult.Screenshot);
                }
            }
            catch (Exception ex) {
                OnError("Errored while trying to parse the confirmationID from the webpage after submitting.", ex, ErrorCategory.Script, ErrorAssignee.Support, 8017);
            }
        }

        [Method("Confirms the agency is still processing the filing or payment.  Instructs the EFP processor try again later.")]
        public void ConfirmProcessingAcknowledgement(int waitMinutes = 0, int waitHours = 0, int waitDays = 0) {
            Log();
            try {
                ScriptResult.ScriptStatus = EFPSubmitter.SubmissionStatus.ProcessingAcknowledgement;
                DateTime nextTryDate = DateTime.UtcNow;
                nextTryDate = nextTryDate.AddDays(waitDays);
                nextTryDate = nextTryDate.AddHours(waitHours);
                nextTryDate = nextTryDate.AddMinutes(waitMinutes);
                ScriptResult.NextTryUTC = nextTryDate;
                SaveScreenshot();
            }
            catch (Exception ex) {
                OnError("Error setting submitter result status to 'ProcessingAcknowledgement' while exectuing the script", ex, ErrorCategory.Script, ErrorAssignee.Support, 8042);
            }
        }

        [Method("Deprecated")]
        public void ConfirmProcessing() {
            Log();
            try {
                ConfirmProcessingAcknowledgement(0, 0, 1);
            }
            catch (Exception ex) {
                OnError("Error setting submitter result status to 'ProcessingAcknowledgement' while exectuing the script", ex, ErrorCategory.Script, ErrorAssignee.Support, 8042);
            }
        }

        [Method("Confirms that we came back to the website after submitting the payment or filing, and the agency has verified that our submission was accepted successfully and formatted correctly without errors.")]
        public void ConfirmAcknowledged() {
            Log();
            try {
                ScriptResult.ScriptStatus = EFPSubmitter.SubmissionStatus.Acknowledged;
                SaveScreenshot();
            }
            catch (Exception ex) {
                OnError("Error setting submitter result status to 'Acknowledged' while exectuing the script", ex, ErrorCategory.Script, ErrorAssignee.Support, 8042);
            }
        }

        [Method("Confirms that the filing or payment was rejected by the agency.")]
        public void ConfirmRejected() {
            Log();
            try {
                ScriptResult.ScriptStatus = EFPSubmitter.SubmissionStatus.Rejected;
                SaveScreenshot();
            }
            catch (Exception ex) {
                OnError("Error setting submitter result status to 'Rejected' while exectuing the script", ex, ErrorCategory.Script, ErrorAssignee.Support, 8042);
            }
        }

        [Method("Confirms that the filing or payment was rejected by the agency.  Use this only for test purposes.")]
        public void ConfirmRejected2(string literalRejectionMessage) {
            Log();
            try {
                ScriptResult.InternalMessage = literalRejectionMessage;
                ScriptResult.Message = literalRejectionMessage;
                ScriptResult.ScriptStatus = EFPSubmitter.SubmissionStatus.Rejected;
                SaveScreenshot();
            }
            catch (Exception ex) {
                OnError("Error setting submitter result status to 'Rejected' while exectuing the script", ex, ErrorCategory.Script, ErrorAssignee.Support, 8042);
            }
        }

        [Method("Deprecated.  Do not use")]
        public void ConfirmBluePrintCheck() {
            Log();
        }

        [Method("When running the blueprint check, if the crawler makes it to this command, the blueprint checker will consider it a positive website check and end the command execution")]
        public void EndIfTest(){
            Log();
            try {
                if (IsATest) {
                    string confirmation = StringExtensions.IsNullOrWhiteSpace(ScriptResult.AgencyConfirmationID) ? Guid.NewGuid().ToString().Substring(0, 10) : ScriptResult.AgencyConfirmationID;
                    if (DataProvider.Submission.SubmissionStatus == SubmissionStatus.Submitting) {
                        ConfirmSubmitted2(confirmation);
                    }
                    else if (DataProvider.Submission.SubmissionStatus == SubmissionStatus.Acknowledging) {
                        ConfirmAcknowledged();
                    }
                    else {
                        ConfirmSubmitted2(confirmation);
                    }
                    ScriptResult.LoginCompleted = true;
                    EndProcessing();
                }
            }
            catch (Exception ex) {
                OnError("Error ending the script executor for a test submission", ex, ErrorCategory.Script, ErrorAssignee.Support, 8044);
            }
        }

        [Method("Validates that a script has successfully logged in by searching the page for the specified text/html.  If the text exists within the page and 'Contains' is selected, then the login will be considered successful.  ")]
        public void ValidateLogin(string value, ContainsOrNotContains containsOrNotContains, ErrorAssignee assignee = ErrorAssignee.Client) {
            Log();
            try {

                bool containsText = If_PageContains(value);

                if ((containsOrNotContains == ContainsOrNotContains.Contains && containsText == true) || (containsOrNotContains == ContainsOrNotContains.DoesNotContain && containsText == false)) {
                    ScriptResult.LoginCompleted = true;
                }
                else {
                    Log("did not validate login");
                    ScriptResult.LoginCompleted = false;
                    OnError("Error while logging in.", null, ErrorCategory.Data, ErrorAssignee.Client, 8020, null);
                }
            }
            catch (Exception ex) {
                ScriptResult.LoginCompleted = false;
                OnError("Error trying to validate the login status", ex, ErrorCategory.Data, ErrorAssignee.Client, 8045, null);
            }
        }

        [Method("Validates that a website is online by searching the page for the specified text/html.  If the text exists within the page and 'Contains' is selected, then the website will be considered online.")]
        public void ValidateWebPage(string value, ContainsOrNotContains containsOrNotContains, ErrorAssignee assignee = ErrorAssignee.Compliance) {
            Log();
            try {

                ScriptResult.WebsiteOnline = false;
                bool containsText = If_PageContains(value);

                if ((containsOrNotContains == ContainsOrNotContains.Contains && containsText == true) || (containsOrNotContains == ContainsOrNotContains.DoesNotContain && containsText == false)) {
                    ScriptResult.WebsiteOnline = true;
                }
                else {
                    try {
                        if (If_URLContains("about:blank")) {
                            ScriptResult.IsAboutBlank = true;
                        }
                    }
                    catch (Exception ex1) {
                        ScriptResult.IsAboutBlank = true;
                        ScriptResult.WebsiteOnline = false;
                        OnError("Error while validating website is online.", ex1, ErrorCategory.Data, ErrorAssignee.Compliance);
                    }
                }
            }
            catch (Exception ex) {
                OnError("Error while validating website is online.", ex, ErrorCategory.Data, ErrorAssignee.Compliance);
            }
        }

        [Method("Validates that the current page contains or does not contain the given text.  If the page does not contain the given text and the ContainsOrNotContains parameter is set to 'Contains', then the script will be stopped with the given error message.  The opposite is true if the ContainsOrNotContains parameter is set to 'DoesNotContain.")]
        public void ValidatePageText(string value, ContainsOrNotContains containsOrNotContains, string ErrorMessage) {
            Log();
            try {

                bool containsText = If_PageContains(value);

                if ((containsOrNotContains == ContainsOrNotContains.Contains && containsText == false) || (containsOrNotContains == ContainsOrNotContains.DoesNotContain && containsText == true)) {
                    OnError(ErrorMessage, null, ErrorCategory.Script, ErrorAssignee.Support, 8018, null);
                }
            }
            catch (Exception ex) {
                OnError("Error trying to validate the page text", ex, ErrorCategory.Data, ErrorAssignee.Client, 8046, null);
            }
        }

        private decimal ConvertDirtyDecimalString(string value) {

            // first we clean the string up
            string cleanValue = value.Replace("$", "").Replace(",", "").Replace(" ", "").Replace("(", "-").Replace(")", "");
            decimal d = 0m;
            decimal.TryParse(cleanValue, out d);
            return d;
        }

        [Method("Validates that the liability that the state says the client owes is equal to the liability that the client is paying.  If the two values are not equal, then a bit will be set and an email will be sent out to the EFP rep at the end of the submission process.  If the 'failIfIncorrect' parameter is set to true, the script will immediately stop execution if the amounts do not match.")]
        public void ValidateLiability(string stateAmount, string clientAmount, bool failIfIncorrect) {
            //Log();
            //try {
            //    ScriptResult.LiabiltyDue = ConvertDirtyDecimalString(stateAmount);
            //    ScriptResult.LiabiltyPaid = ConvertDirtyDecimalString(clientAmount);

            //    if (ScriptResult.LiabilityIncorrect && failIfIncorrect) {
            //        OnError(string.Format("The liabilty amount that the client submitted ({0}) was not equal to the amount on the states website ({1}), and we could not make the exact payment the client requested.", clientAmount, stateAmount), null, ErrorCategory.Script, ErrorAssignee.Support, 8047, null);
            //    }
            //}
            //catch (Exception ex) {
            //    OnError("Error validating the client liability matches the state's liability", ex, ErrorCategory.Script, ErrorAssignee.Support, 8048, null);
            //}
        }

        # region common actions
        [Method("Clicks an element of the specified type given its unique idtype and associated value.  NoWait is an optional parameter that" 
         +" tells the processor not to wait for the page to load.  If 'RadioButton' is selected, the groupName must also be specified.")]
        public void Click(BrowserDriver.ElementType elementType, ElementAttributes idType, string idValue, bool noWait = false, string groupName = null, int elementIndex = 0) {
            Log();
            try {

                Log("Getting click constraint");
               
                    Log("Getting element");
                    Browser.FindElement(elementType, idType, idValue);

                if (!Browser.Element.IsNull){
                    Browser.Element.Click();
                }
                else {
                    OnError("Could not find a " + elementType + " with constraint " + idType + ": " + idValue, null, ErrorCategory.Script, ErrorAssignee.Development, 8049, null);
                }
            }
            catch (Exception ex) {
                OnError("Error trying to click an html element on the current page.", ex, ErrorCategory.Script, ErrorAssignee.Development, 8049, null);
            }
        }


        [Method("Add a cookie to the browser. Useful when dealing with websites that require access codes so we don't appear as a new machine.")]
        public void AddCookie(string cookieName, string cookieValue, string domain, string path) {
            Browser.CookieManager.AddCookie(new Cookies(cookieName, cookieValue, domain, path));
        }


        [Method("Clicks a javascript virtual element of the specified type given its unique idtype and associated value.")]
        public void ClickJavascriptElement(BrowserDriver.ElementType elementType, ElementAttributes idType, string idValue, bool noWait = false, string groupName = null, int elementIndex = 0) {
            Log();
            try {

                Log("Getting click constraint");

                Log("Getting element");
                Browser.FindElement(elementType, idType, idValue);

                if (!Browser.Element.IsNull) {
                    Browser.ClickVirtualElement(Browser.Element);                   
                }
                else {
                    OnError("Could not find a " + elementType + " with constraint " + idType + ": " + idValue, null, ErrorCategory.Script, ErrorAssignee.Development, 8049, null);
                }
            }
            catch (Exception ex) {
                OnError("Error trying to click a javascript element on the current page.", ex, ErrorCategory.Script, ErrorAssignee.Development, 8049, null);
            }
        }  


        [Method("Reloads the DOM for the specified element.  Use this for persistent timeout errors on specific elements.")]
        public void RefreshElement(Element value) {
            Log();
            try {
                Browser.Refresh(value);
            }
            catch { }
        }

        [Method("Refreshes the webpage.")]
        public void Refresh() {
            Log();
            try {
                Browser.Refresh();
            }
            catch { }
        }

        private bool UseXY = true;
        [Method("Sets whether or not an image click should send its X and Y coordinates.  Some websites require that they be sent, some websites do not. If you have trouble with Image clicks, use this and see if it corrects the problem.")]
        public void SetImageXY(OnOrOff onOrOff) {
            Log();
            try {
                if (onOrOff == OnOrOff.On) {
                    UseXY = true;
                }
                else if (onOrOff == OnOrOff.Off) {
                    UseXY = false;
                }
            }
            catch (Exception ex) {
                OnError(null, ex);
            }
        }

        [Method("Selects ann item from a drop down list with the text specified by 'itemTextToSelect'")]
        public void SelectFromListBox(BrowserDriver.ElementAttributes idType, string idValue, ElementAttributes findBy, string itemTextToSelect) {
            Log();
            try {
                Browser.SelectFromListBox(idType, idValue, findBy, itemTextToSelect);
            }
            catch (Exception ex) {
                OnError("Error trying to select an item from a list box on the page", ex, ErrorCategory.Script, ErrorAssignee.Support, 8051, null);
            }
        }

        [Method("Physically sets the text of the in memory textbox programatically")]
        public void SetTextCurrentElement(string textToInsert) {
            Log();
            try {
                if (textToInsert == null || textToInsert == "NULL") {
                    OnError("I tried to insert text into a webpage, but the text was null.", null, ErrorCategory.Script, ErrorAssignee.Support, 8052, null);
                }
                else {
                    Log("Checking existance");
                    if (!Browser.Element.IsNull) {
                        if (!Browser.Element.Enabled) {
                            try {
                                Browser.SetAttribute(Browser.Element, "disabled", "false");
                            }
                            catch { }
                        }

                        Log("Setting Value");
                        Browser.Element.SendKeys(textToInsert);
                        Browser.Element.Blur();

                        Log("Value set");
                    }
                    else {
                        OnError("Could not find a textField to set text on.", null, ErrorCategory.Script, ErrorAssignee.Support, 8052, null);
                    }
                }

            }
            catch (Exception ex) {
                OnError("Error trying to set the text of a specified element on the page", ex, ErrorCategory.Script, ErrorAssignee.Support, 8052, null);
            }
        }


        [Method("Physically sets the text of the specified textbox programatically")]
        public void SetText(ElementAttributes idType, string idValue, string textToInsert) {
            Log();
            try {

                if (textToInsert == null || textToInsert == "NULL") {
                    OnError("I tried to insert text into a webpage, but the text was null.", null, ErrorCategory.Script, ErrorAssignee.Support, 8052, null);
                }
                else {
                    Log("Getting setText Constraint");
                    Element input = Browser.FindElement(BrowserDriver.ElementType.Input, idType, idValue);
                    input.Focus();
                    input = Browser.FindElement(BrowserDriver.ElementType.Input, idType, idValue);

                    Log("Checking existance");
                    if (!input.IsNull) {
                        if (!input.Enabled) {
                            try {
                                Browser.SetAttribute(input, "disabled", "false");
                            }
                            catch { }
                        }

                        Log("Setting Value");
                            input.SendKeys(textToInsert);
                            input.Blur();
                     
                        Log("Value set");
                    }
                    else {
                        OnError("Could not find a textField with with " + idType + " equal to " + idValue, null, ErrorCategory.Script, ErrorAssignee.Support, 8052, null);
                    }
                }
            }
            catch (Exception ex) {
                OnError("Error trying to set the text of a specified element on the page", ex,  ErrorCategory.Script, ErrorAssignee.Support, 8052, null);
            }
        }

        [Method("Physically sets the text of the specified textbox programatically")]
        public void TypeText(ElementAttributes idType, string idValue, string textToInsert, bool append = false, bool sendTabKey = false) {
        //    Log();
        //    TypeTextWithBlurOption(idType, idValue, textToInsert, append, sendTabKey, false);
        }

        private void TypeTextWithBlurOption(ElementAttributes idType, string idValue, string textToInsert, bool append = false, bool sendTabKey = false, bool ignoreClearBlur = false) {
        //    Log();
        //    try {

        //        if (textToInsert == null || textToInsert == "NULL") {
        //            OnError("I tried to insert text into a webpage, but the text was null.", null, ErrorCategory.Script, ErrorAssignee.Support, 8052, null);
        //        }
        //        Constraint constraint = GetElementConstraint(idType, idValue, true);

        //        TextField textField = IE.TextField(constraint);
        //        if (textField.Exists) {
        //            if (append == false) {
        //                if (ignoreClearBlur) {
        //                    textField.TypeTextAction.ClearIgnoreBlur();
        //                }
        //                else {
        //                    textField.TypeTextAction.Clear();
        //                }
        //            }
        //            textField.TypeTextAction.AppendText(sendTabKey ? textToInsert + "\t" : textToInsert);

        //            if (DataProvider.Data["Jurisdiction"].ToString() == "US-NY") {
        //                WatiN.Core.UtilityClasses.UtilityClass.TryActionIgnoreException(textField.Blur);
        //            }

        //        }
        //        else {
        //            OnError("Could not find a textField with with " + idType + " equal to " + idValue, null, ErrorCategory.Script, ErrorAssignee.Support, 8052, null);
        //        }
        //    }
        //    catch (Exception ex) {
        //        OnError("Error trying to set the text of a specified element on the page", ex, ErrorCategory.Script, ErrorAssignee.Support, 8052, null);
        //    }
        }

        [Method("Fills in a file upload box with the current EFile obtained from the crawler's dataprovider.  If the fileName is specified, then it will override the file name set in the spec.  If the submission has two efiles, then the specific specID needs to be specified.")]
        public void SelectUploadFile(ElementAttributes idType, string idValue, string fileName = null, int specID = 0, bool shouldZip = false) {
            Log();
            try {

                string tempFile = GetFileLocation(fileName, specID, shouldZip);
                try {
                    Browser.FileUpload(BrowserDriver.ElementType.Input, idType, idValue, tempFile);
                }


                catch (Exception ex) {
                    OnError("Could not set the file to upload to the filename selected. \n\n " + ex.ToString(), ex, ErrorCategory.Script, ErrorAssignee.Development, 8053, null);
                }
            }
            catch (Exception ex) {
                OnError("Error setting a file upload dialog box with to the correct value", ex, ErrorCategory.Script, ErrorAssignee.Development, 8053, null);
            }
        }

        [Method("Gets the path to the file that should be uploaded.  If the fileName is specified, then it will override the file name set in the spec.  If the submission has two efiles, then the specific specID needs to be specified.")]
        public string GetFileLocation(string fileName = null, int specID = 0, bool shouldZip = false) {
            Log();
            try {
                Greenshades.EFP.EFPSubmitter.SubmissionFile file = DataProvider.Files.Where(f => f.FileType == "EFile").ToList()[0];
                if (specID != 0) {
                    file = DataProvider.Files.First(f => f.FileType == "EFile" && f.SpecID == specID);
                }

                string fixedfileName = file.FileName;
                foreach(char c in Path.GetInvalidPathChars().Union(Path.GetInvalidFileNameChars())) {
                    fixedfileName = fixedfileName.Replace(c.ToString(), "");
                }

                string defaultFileName = string.Format(@"{0}\{1}\{2}", Environment.GetEnvironmentVariable("TEMP"), Guid.NewGuid(), fixedfileName);
                string userFileName = string.Format(@"{0}\{1}\{2}{3}", Environment.GetEnvironmentVariable("TEMP"), Guid.NewGuid(), fileName, Path.GetExtension(fixedfileName));

                // if the script creator specifies a filename, then use that instead of the default file name set in the spec.
                // if we need to zip the file, then make the zip file by replacing the extension with .zip
                string unZippedFileName = defaultFileName;
                if (string.IsNullOrEmpty(fileName) == false && shouldZip == false) {
                    unZippedFileName = userFileName;
                }

                Directory.CreateDirectory(Path.GetDirectoryName(unZippedFileName));
                File.WriteAllBytes(unZippedFileName, file.FileBytes.ToArray());

                string zippedFileName = string.Empty;
                if (shouldZip) {

                    byte[] zippedBytes = Greenshades.Utils.ZipUtils.ZipFile(file.FileBytes.ToArray(), Path.GetFileName(unZippedFileName));
                    string extension = Path.GetExtension(unZippedFileName);
                    zippedFileName = unZippedFileName.Replace(extension, ".zip");

                    File.WriteAllBytes(zippedFileName, zippedBytes);
                }

                return zippedFileName == string.Empty ? unZippedFileName : zippedFileName;
            }
            catch (Exception ex) {
                OnError("Failed attempting to obtain a file location when submitting an efile.", ex, ErrorCategory.Script, ErrorAssignee.Development, 8063);
            }
            return null;
        }

        [Method("Responds to a download file dialog by clicking the save option and saving the file.  The file will be saved in the script-result receipts.")]
        public void DownloadFile(string fileNameToSaveAs, ReceiptType fileType) {
            Log();
            try {

                // this is so that if we click on a pdf file, the download dialog box is shown instead of opening the pdf automatically
                
                if (Registry.ClassesRoot.OpenSubKey(".pdf") != null) {
                    Registry.ClassesRoot.DeleteSubKeyTree(".pdf");
                };

                string tempFile = Browser.DownloadPath;

                if (Browser.CurrentBrowserType == BrowserType.Chrome) {
                    tempFile = string.Format(@"{0}\", Browser.DownloadPath);
                    Directory.CreateDirectory(Path.GetDirectoryName(tempFile));
                }
                else {
                    tempFile = Browser.DownloadPath;
                }
                

                Browser.DialogHandler.SaveFileDialog(tempFile + fileNameToSaveAs);

                if (File.Exists(tempFile + fileNameToSaveAs)) {
                    AddReceipt(tempFile, fileType);
                                         
                    File.Delete(tempFile + fileNameToSaveAs);
                }
                else {
                    OnError("Error when trying to add a downloaded file to the receipts folder.  File does not exist, file path: " + tempFile + fileNameToSaveAs);
                }

            }
            catch (Exception ex) {
                OnError("Error downloading file from website during submission script", ex, ErrorCategory.Script, ErrorAssignee.Development, 8065, null);
            }
        }

        [Method("Navigates the browser to a specific webpage")]
        public void GoToUrl(string url, bool noWait = false) {
            Log();
            try {
                Browser.GoToUrl(url);
            }
            catch (Exception ex) {
                OnError("Error navigating to url " + url, ex, ErrorCategory.Script, ErrorAssignee.Development, 8064, null);
            }
        }

        [Method("Saves the current view of the webpage to a temporary file that can later be accessed")]
        public void SaveScreenshot() {
            Log();
            try {
                    // save the window state
                    Log("Taking screenshot");
                    byte[] fileBytes = Browser.TakeScreenshot();
                    ScriptResult.Screenshot = fileBytes; 
                    Log("Screenshot complete");
            }
            catch (Exception ex) {
                ex.ToString();
            }
        }

        [Method("Returns the first matching text from the current page's html from the given REGEX. If no match is found, the empty string is returned.")]
        public string RegexMatch(string value) {
            Log();
            try {
                string currentHTML = Browser.Html;

                Match matchResult = Regex.Match(currentHTML, value);
                if (matchResult.Success) {
                    return matchResult.Value;
                }

                return string.Empty;
            }
            catch (Exception ex) {
                OnError("Error obtaining regex match from script", ex, ErrorCategory.Script, ErrorAssignee.Development);
                return string.Empty;
            }
        }
        

        [Method("Checks if the difference between the payment amount and the website payment amount are within the payment threshold, use Website amount as payment amount if validation passes.")]
        public void ValidatePaymentThreshold(decimal paymentAmount, decimal websitePaymentAmount, decimal paymentThreshold) {
            decimal difference = Math.Abs(paymentAmount - websitePaymentAmount);

            if (difference > paymentThreshold) {
                End("Payment Amount was outside of the set Payment Threshold on this payment.", true, ErrorCategory.Data, ErrorAssignee.Support, 8602, null);
            }

            this.ScriptResult.DifferingAmounts.DifferentAmountPaid = true;
            this.ScriptResult.DifferingAmounts.OriginalAmount = paymentAmount;
            this.ScriptResult.DifferingAmounts.AmountPaid = websitePaymentAmount;
            this.ScriptResult.DifferingAmounts.ThresholdAmount = paymentThreshold;
            return;
        }

        [Method("Ends the current script's execution errormessage specified in this elements value node.  If the endWithError option is set to false, then no error will be generated.")]
        public void End(string value = null, bool endWithError = true, ErrorCategory category = ErrorCategory.Script, ErrorAssignee assignee = ErrorAssignee.Development, int ErrorID = 8001, string AgencyErrorID = null) {
            Log();
            try {
                if (endWithError) {
                    OnError(value, null, category, assignee, ErrorID, AgencyErrorID);
                }
                else {                   
                    if (ScriptResult.Screenshot == null) {
                        SaveScreenshot();
                    }

                    EndProcessing();
                }
            }
            catch (Exception ex) {
                OnError("Error ending a script execution when processing and 'End' node", ex, ErrorCategory.Script, ErrorAssignee.Development, 8067, null);
            }
        }

        # endregion common actions

        # region uncommon actions
        [Method("Attaches the browser instance to a new window with the given constraints")]
        public void AttachToWindow(WindowSearchType idType, string idValue) {
            Log();
            try {
                //if (idType != WindowSearchType.URL && idType != WindowSearchType.Title) {
                //    throw new Exception("Only URL or Title can be used for the idType when attaching to a new window.");
                //}

                Browser.AttachTo(WindowType.Window, idType, idValue);
            }
            catch (Exception ex) {
                OnError("Error attaching to an IE window.", ex, ErrorCategory.Script, ErrorAssignee.Development, 8066, null);
            }
        }

        [Method("Attaches the browser instance to a specific frame within the page.")]
        public void AttachToFrame(ElementAttributes idType, string idValue) {
            Log();
            try {
                try {
                    Browser.AttachToIframe(idType, idValue);
                }
                catch (Exception ex) {
                    OnError("Could not find and attach to Frame with " + idType + " equal to " + idValue, ex, ErrorCategory.Script, ErrorAssignee.Development, 8068, null);
                }
            }
            catch (Exception ex) {
                OnError("Error Attaching to html frame", ex, ErrorCategory.Script, ErrorAssignee.Development, 8068, null);
            }
        }

        [Method("Detaches from the current frame and reattaches to the main browser html window")]
        public void DetachFromFrame(bool closeWindow) {
            Log();
            try {
                if (closeWindow) {
                    Browser.DetachToMain();
                }
                else{
                    Browser.DetachToMainWithoutClosing();
                }
            }
            catch (Exception ex) {
                OnError("Error Detaching from html frame", ex, ErrorCategory.Script, ErrorAssignee.Development, 8068, null);
            }
        }

        [Method("Fills in a windows login dialog with the given credentials.  If multiple applications are open, this command is susceptable to failure.")]
        public void FillLogonDialog(string username, string password, LogonButton buttonToClick) {
            Log();
            ///this guy needs to be converted and added to the BrowserDriver api to provide a simple method of attaching to a logon dialog box.
            try {
                Process[] processes = Process.GetProcessesByName("iexplore");
                foreach (Process proc in processes) {

                    if (proc.MainWindowHandle == IntPtr.Zero) {
                        continue;
                    }

                    Condition edit = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                    Condition window = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
                    Condition button = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                    Condition enabled = new PropertyCondition(AutomationElement.IsEnabledProperty, true);

                    AutomationElement main = AutomationElement.FromHandle(proc.MainWindowHandle);
                    AutomationElementCollection windows = main.FindAll(TreeScope.Children, new AndCondition(window, enabled));
                    foreach (AutomationElement childWindow in windows) {

                        if (childWindow.Current.ClassName == "#32770") {

                            AutomationElementCollection edits = childWindow.FindAll(TreeScope.Descendants, new AndCondition(edit, enabled));
                            foreach (AutomationElement childEdit in edits) {

                                if (childEdit.Current.Name == "User name") {
                                    SetAutomationText(childEdit, username);
                                }

                                if (childEdit.Current.Name == "Password") {
                                    SetAutomationText(childEdit, password);
                                }
                            }

                            AutomationElementCollection buttons = childWindow.FindAll(TreeScope.Descendants, new AndCondition(button, enabled));
                            foreach (AutomationElement childButton in buttons) {


                                if (childButton.Current.Name == buttonToClick.ToString()) {
                                    ClickAutomationElement(childButton);
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) {
                OnError("Error filing in a windows authentication login dialog box", ex, ErrorCategory.Script, ErrorAssignee.Development, 8069, null);
            }
        }

        [Method("Clicks a button on a javascript popup dialog box")]
        public void ClickDialogButton(DialogType dialogType, DialogButtons dialogButton, bool noWait = false) {
            Log();
            try {
                
                System.Threading.Thread.Sleep(1000);
                Browser.DialogHandler.AlertBoxClose(dialogButton);
            }
            catch (Exception ex) {
                OnError("Error clicking a dialog button", ex, ErrorCategory.Script, ErrorAssignee.Development, 8070, null);
            }
        }

        [Method("When set to on, all popup dialog windows will automatically close.")]
        public void SetAutoCloseDialogs(OnOrOff onOrOff) {
            Log();
            try {
                if (onOrOff == OnOrOff.On) {
                    Browser.DialogHandler.AutoCloseDialogBox(true);
                }
                else {
                    Browser.DialogHandler.AutoCloseDialogBox(false);
                }
            }
            catch (Exception ex) {
                OnError("Error changing the status of the AutoCloseDialogs property on the script executor", ex, ErrorCategory.Script, ErrorAssignee.Development, 8071, null);
            }
        }

        [Method("Pauses the execution for the given period of time in milliseconds")]
        public void WaitMilliSeconds(int millisecondsToWait) {
            Log();
            try {
                System.Threading.Thread.Sleep(millisecondsToWait);
            }
            catch (Exception ex) {
                OnError("Error waiting a specified time amount while running the script executor", ex, ErrorCategory.Script, ErrorAssignee.Development, 8073, null);
            }
        }

        [Method("Closes the tubmleweed upload dialog box")]
        public void CloseTumbleweedUploadConfirmation(string fileNameUploaded, bool ignoreCheck = false) {
            Log();
            try {

                System.Threading.Thread.Sleep(5000);
                Browser.DialogHandler.CloseDialogBox("SecureTranport transfer status", "OK");

                // ok if we got this far, that means we diddnt find the close dialog box.  For some reason we dont always get the dialog box, so need to check here if the page shows that the file was uploaded.
                System.Threading.Thread.Sleep(2000);

                //ok if there was no dialog, then the page should show our file.
                bool success = ignoreCheck || If_PageContains(fileNameUploaded);

                if (!success) {
                    OnError("I could not find the OK button for the tumbleweed upload dialog, and the current page does not show my file.", null, ErrorCategory.Script, ErrorAssignee.Development, 8072, null);
                }
            }
            catch (Exception ex) {
                OnError("Error closing the tumbleweed dialog box", ex, ErrorCategory.Script, ErrorAssignee.Development, 8072, null);
            }
        }

        [Method("Pauses the execution until the specified element exists up to the maximum milliseconds to wait.  ")]
        public void WaitForElement(BrowserDriver.ElementType elementTag, ElementAttributes idType, string idValue, int secondsToWait, string inputType = null) {
            Log();
            try {

                string[] inputTypes = null;
                if (inputType != null) {
                    inputTypes = new string[] { inputType };
                }

                try {
                    Browser.WaitForElementToExist(elementTag, idType, idValue, secondsToWait);
                }
                catch (Exception ex) {            
                    OnError("Error waiting for element to exist during call to WaitForElement(" + elementTag + ", " + idType + ", " + idValue + ", " + secondsToWait + ")", null, ErrorCategory.Script, ErrorAssignee.Development, 8074, null);
                }
            }
            catch (Exception ex) {
                OnError("Error waiting for an element to show up on a webpage", ex, ErrorCategory.Script, ErrorAssignee.Development, 8074, null);
            }
        }

        [Method("Waits until the page is completely loaded or done processing.  Normally used after a ClickNoWait to avoid a timeout.")]
        public void WaitForComplete(int secondsToWait) {
            Log();
            try {
                Browser.WaitForComplete(secondsToWait);
            }
            catch (Exception ex) {
                OnError("Error waiting for the webpage to complete loading", ex, ErrorCategory.Script, ErrorAssignee.Development, 8075, null);
            }
        }

        [Method("Executes a string of javascript code on the current instance of the browser")]
        public void ExecuteJavaScript(string script) {
            Log();
            try {
                Browser.ExecuteJavascript(script);
            }
            catch (Exception ex) {
                OnError("Error running script", ex, ErrorCategory.Script, ErrorAssignee.Development, 8076, null);
            }
        }

        [Method("Returns to the previous page")]
        public void Back() {
            Log();
            try {
                Browser.Back();
            }
            catch (Exception ex) {
                OnError("Error going back to the previoius webpage", ex, ErrorCategory.Script, ErrorAssignee.Development, 8077, null);
            }
        }

        # endregion uncommon actions

        #region validation stuff

        [Method("Validates that a submission parameter has failed, succeeded, or is in an unknown status")]
        public void ValidateParameter(string parameterName, ParameterStatus status, string agencyValue, string clientValue) {
            Log();
            try {
                VerificationParameter param = ScriptResult.VerificationParameters.Single(p => p.ParameterName == parameterName);
                param.ParameterStatus = (EFP.ParameterStatus)Enum.Parse(typeof(EFP.ParameterStatus), status.ToString());
                if (!string.IsNullOrEmpty(agencyValue)) {
                    param.AgencyValue = agencyValue;
                }
                if (!string.IsNullOrEmpty(clientValue)) {
                    param.ClientValue = clientValue;
                }
            }
            catch (Exception ex) {
                OnError("Error updating validation Parameter", ex, ErrorCategory.Script, ErrorAssignee.Development);
            }
        }

        [Method("Validates that a submission parameter has failed, succeeded, or is in an unknown status")]
        public void EndIfValidationParameterFailed() {
            Log();
            try {
                string errorMessage = string.Empty;
                if (ScriptResult.VerificationParameters.Any(p => p.ParameterStatus == ParameterStatus.Failed)) {
                    foreach (var param in ScriptResult.VerificationParameters.Where(s => s.ParameterStatus == ParameterStatus.Failed)) {
                        errorMessage += string.Format("{0}:\t\t\t{1}", param.ParameterName, param.ParameterStatus.ToString());
                    }

                    OnError("Registration parameter validation failed: \r\n" + errorMessage, null, ErrorCategory.Data, ErrorAssignee.Client);
                }
            }
            catch (Exception ex) {
                OnError("Error checking validation Parameter", ex, ErrorCategory.Script, ErrorAssignee.Development);
            }
        }

        #endregion validation stuff

        [Method("This takes a comma separated list of IE version numbers (8,10) that this script supports.  If the machine that the script is running on does not have a supoprted version of IE installed, the script will end and the submission will return to the queue. ")]
        public void SetSupportedIEVersions(string value) {
            try{
                if (value != "Selenium") {
                    ScriptResult.ScriptStatus = SubmissionStatus.VersionNotSupported;
                    ScriptResult.VersionsSupported = value.ToString();
                    End("Machine does not support Watin scripts", false);
                }
            }
            catch (Exception ex) {
                OnError("Error finding the correct browser version", ex, ErrorCategory.Script, ErrorAssignee.Development);
            }
        }

        [Method("Specifies how long the WatiN default timeout should be. Put in the section before the login group.")]
        public void SetDefaultTimeout(int value) {
        }

        [Method("Determines which browser type the script will use.  If left out will default to Chrome.")]
        public void SetBrowserType(BrowserType browserType) {
        }

        #endregion Available Commands
    }
}


