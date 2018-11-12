////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Alivia Houdek 
// 08.20.2018 
// Universal Returns Automation 
// Run on every Saturday AND every first of the month

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AHK_Universal_Returns
{
    class ClientGroup
    {
        public int cPrimaryId;
        public string cName;
        public string cGroup;
        public string cPath;
        public string cReturnEmail;
        public string cSettleEmail;
        public string cBankoEmail;
        public string cLastRun;
        public string cPassword;
        public string cFrequency;
        public string cBreakOut;
        public string cItemizedSpecial;
        public string cPersonAdded;
        public string cDateAdded;
        public string cPersonLastUpdated;
        public string cLastUpdatedDate;
        public string cComments;
        public string cDistinctReturns;
        public string cCustomSplit;
        public string cSeparateByType;
        public string cSeparateByStatus;
        public string cSeparateByClient;
        public string cHeaders;

        // Standard client group with no custom splitting (could have different headers)
        public ClientGroup(int inId, string inName, string inGroup, string inPath, string inReturnEmail, string inSettleEmail,
            string inBankoEmail, string inLastRun, string inPassword, string inFrequency, string inBreakOut, string inItemizedSpecial,
            string inPersonAdded, string inDateAdded, string inPersonLastUpdated, string inLastUpdatedDate, string inComments, string inComprehensiveReturns, string inCustomSplit, string inHeaders)
        {
            cPrimaryId = inId;
            cName = inName.Trim();
            cGroup = inGroup.Trim().ToUpper();
            cPath = inPath.Trim().ToUpper();
            cReturnEmail = inReturnEmail.Trim().ToUpper();
            cSettleEmail = inSettleEmail.Trim().ToUpper();
            cBankoEmail = inBankoEmail.Trim().ToUpper();
            cLastRun = inLastRun.Trim().ToUpper();
            cPassword = inPassword.Trim();
            cFrequency = inFrequency.Trim().ToUpper();
            cBreakOut = inBreakOut.Trim().ToUpper();
            cItemizedSpecial = inItemizedSpecial.Trim().ToUpper();
            cPersonAdded = inPersonAdded.Trim().ToUpper();
            cDateAdded = inDateAdded.Trim().ToUpper();
            cPersonLastUpdated = inPersonLastUpdated.Trim().ToUpper();
            cLastUpdatedDate = inLastUpdatedDate.Trim().ToUpper();
            cComments = inComments.Trim().ToUpper();
            cDistinctReturns = inComprehensiveReturns.Trim().ToUpper();
            cCustomSplit = inCustomSplit.Trim().ToUpper();
            cSeparateByType = "";
            cSeparateByStatus = "";
            cSeparateByClient = "";
            cHeaders = inHeaders;
        }

        // Client with custom splitting by files or sheets by status code only
        public ClientGroup(int inId, string inName, string inGroup, string inPath, string inReturnEmail, string inSettleEmail,
           string inBankoEmail, string inLastRun, string inPassword, string inFrequency, string inBreakOut, string inItemizedSpecial,
           string inPersonAdded, string inDateAdded, string inPersonLastUpdated, string inLastUpdatedDate, string inComments, string inComprehensiveReturns, string inCustomSplit, string inSeparateByType, string inSeparateByStatus, string inHeaders)
        {
            cPrimaryId = inId;
            cName = inName.Trim();
            cGroup = inGroup.Trim().ToUpper();
            cPath = inPath.Trim().ToUpper();
            cReturnEmail = inReturnEmail.Trim().ToUpper();
            cSettleEmail = inSettleEmail.Trim().ToUpper();
            cBankoEmail = inBankoEmail.Trim().ToUpper();
            cLastRun = inLastRun.Trim().ToUpper();
            cPassword = inPassword.Trim();
            cFrequency = inFrequency.Trim().ToUpper();
            cBreakOut = inBreakOut.Trim().ToUpper();
            cItemizedSpecial = inItemizedSpecial.Trim().ToUpper();
            cPersonAdded = inPersonAdded.Trim().ToUpper();
            cDateAdded = inDateAdded.Trim().ToUpper();
            cPersonLastUpdated = inPersonLastUpdated.Trim().ToUpper();
            cLastUpdatedDate = inLastUpdatedDate.Trim().ToUpper();
            cComments = inComments.Trim().ToUpper();
            cDistinctReturns = inComprehensiveReturns.Trim().ToUpper();
            cCustomSplit = inCustomSplit.Trim().ToUpper();
            cSeparateByType = inSeparateByType.Trim().ToUpper();
            cSeparateByStatus = inSeparateByStatus.Trim().ToUpper();
            cSeparateByClient = "";
            cHeaders = inHeaders;
        }

        // Client group with custom splitting of files or sheets by status code and/or client code
        public ClientGroup(int inId, string inName, string inGroup, string inPath, string inReturnEmail, string inSettleEmail,
            string inBankoEmail, string inLastRun, string inPassword, string inFrequency, string inBreakOut, string inItemizedSpecial,
            string inPersonAdded, string inDateAdded, string inPersonLastUpdated, string inLastUpdatedDate, string inComments, string inComprehensiveReturns, string inCustomSplit, string inSeparateByType, string inSeparateByStatus, string inSeparateByClient, string inHeaders)
        {
            cPrimaryId = inId;
            cName = inName.Trim();
            cGroup = inGroup.Trim().ToUpper();
            cPath = inPath.Trim().ToUpper();
            cReturnEmail = inReturnEmail.Trim().ToUpper();
            cSettleEmail = inSettleEmail.Trim().ToUpper();
            cBankoEmail = inBankoEmail.Trim().ToUpper();
            cLastRun = inLastRun.Trim().ToUpper();
            cPassword = inPassword.Trim();
            cFrequency = inFrequency.Trim().ToUpper();
            cBreakOut = inBreakOut.Trim().ToUpper();
            cItemizedSpecial = inItemizedSpecial.Trim().ToUpper();
            cPersonAdded = inPersonAdded.Trim().ToUpper();
            cDateAdded = inDateAdded.Trim().ToUpper();
            cPersonLastUpdated = inPersonLastUpdated.Trim().ToUpper();
            cLastUpdatedDate = inLastUpdatedDate.Trim().ToUpper();
            cComments = inComments.Trim().ToUpper();
            cDistinctReturns = inComprehensiveReturns.Trim().ToUpper();
            cCustomSplit = inCustomSplit.Trim().ToUpper();
            cSeparateByType = inSeparateByType.Trim().ToUpper();
            cSeparateByStatus = inSeparateByStatus.Trim().ToUpper();
            cSeparateByClient = inSeparateByClient.Trim().ToUpper();
            cHeaders = inHeaders.Trim().ToUpper();
        }

        public static string GetCredName(ClientGroup creditorObj)
        {
            String credName = creditorObj.cName;
            /*
            // name formatting 
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            string lower = credName.ToLower(); // whole name set to lowercase
            string[] nameSplit = lower.Split(' ', '-', '&');
            string titledWord = "";
            StringBuilder sb = new StringBuilder();
            int splitCounter = 1;
            foreach (string word in nameSplit)
            {
                titledWord = textInfo.ToTitleCase(word); // capitalizes first letter of word
                if (nameSplit.Length == 1 || nameSplit.Length == splitCounter) {
                    sb.Append(titledWord);
                }
                else if (nameSplit.Length > splitCounter)
                {
                    sb.Append(titledWord + " ");
                    splitCounter++;
                }
                else
                {
                    throw new Exception("Error in setting creditor name in ClientGroup.cs");
                }
            }
            credName = sb.ToString(); */
            return credName;
        }

        public static string GetCredGroup(ClientGroup creditorObj)
        {
            String credGroup = creditorObj.cGroup;
            return credGroup.ToUpper();
        }

        public static string GetCredPath(ClientGroup creditorObj)
        {
            String credPath = creditorObj.cPath;
            if (!(credPath.EndsWith(@"\")))
            {
                credPath = credPath + @"\";
            }
            return credPath;
        }

        public static string[] GetCredReturnEmails(ClientGroup creditorObj)
        {
            String temp = creditorObj.cReturnEmail;
            String[] credReturnEmail = temp.ToUpper().Split(';');
            return credReturnEmail;
        }

        public static string[] GetCredSettleEmails(ClientGroup creditorObj)
        {
            String temp = creditorObj.cSettleEmail;
            String[] credSettleEmail = temp.ToUpper().Split(';');
            return credSettleEmail;
        }

        public static string[] GetCredBankoEmails(ClientGroup creditorObj)
        {
            String temp = creditorObj.cReturnEmail;
            String[] credReturnEmail = temp.ToUpper().Split(';');
            return credReturnEmail;
        }

        public static string GetCredLastRunDate(ClientGroup creditorObj)
        {
            String credRun = creditorObj.cLastRun;
            return credRun.ToUpper();
        }

        public static string GetCredPass(ClientGroup creditorObj)
        {
            String credPass = creditorObj.cPassword;
            return credPass;
        }

        public static string GetCredFrequency(ClientGroup creditorObj)
        {
            String credFreq = creditorObj.cFrequency;
            return credFreq.ToUpper();
        }

        public static string GetCredBreakout(ClientGroup creditorObj)
        {
            String credBO = creditorObj.cBreakOut;
            return credBO.ToUpper() ;
        }

        public static string GetCredItemized(ClientGroup creditorObj)
        {
            String credItemiz = creditorObj.cItemizedSpecial;
            return credItemiz.ToUpper();
        }

        public static string GetCredPersonAdded(ClientGroup creditorObj)
        {
            String credPersonAdded = creditorObj.cPersonAdded;
            return credPersonAdded.ToUpper();
        }

        public static string GetCredDateAdded(ClientGroup creditorObj)
        {
            String credDtAdd = creditorObj.cDateAdded;
            return credDtAdd.ToUpper();
        }

        public static string GetCredPersonLastUpdated(ClientGroup creditorObj)
        {
            String credPrsnUp = creditorObj.cPersonLastUpdated;
            return credPrsnUp.ToUpper();
        }

        public static string GetCredLastUpDate(ClientGroup creditorObj)
        {
            String credLastUp = creditorObj.cLastUpdatedDate;
            return credLastUp.ToUpper();
        }

        public static string GetCredDistinctReturns(ClientGroup creditorObj)
        {
            String credCompreReturns = creditorObj.cDistinctReturns;
            return credCompreReturns.ToUpper();
        }

        public static string GetCredCustomSplit(ClientGroup creditorObj)
        {
            String credCustomSplit = creditorObj.cCustomSplit;
            return credCustomSplit.ToUpper();
        }

        public static string GetCredSplitType(ClientGroup creditorObj)
        {
            String credSplitType = creditorObj.cSeparateByType;
            return credSplitType.ToUpper();
        }

        public static string[] GetCredSplitStatus(ClientGroup creditorObj)
        {
            String stats = creditorObj.cSeparateByStatus;
            string[] credSplitStat = stats.ToUpper().Split(';');
            return credSplitStat;
        }

        public static string[] GetCredSplitClient(ClientGroup creditorObj)
        {
            String clients = creditorObj.cSeparateByClient;
            string[] credSplitClient = clients.ToUpper().Split(';');
            return credSplitClient;
        }

        public static string[] GetCredHeaders(ClientGroup creditorObj)
        {
            String temp = creditorObj.cHeaders;
            String[] credHeaders = temp.Split(';');
            return credHeaders;
        }
    }
}
