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
    class ReturnAccount
    {
        public string acctStatus;
        public string acctCreditorGroup;
        public string acctDebtorNumber;
        public string acctAcctNumber;
        public string acctGuarantorName;
        public string acctPatientName;
        public string acctReturnDescription;
        public string acctReturnAmount;
        public string acctAmountListed;
        public string acctReturnDate;
        public string acctLastPayDate;
        public string acctLastPayAmount;
        public string acctCreditorNumber;
        public string acctDateOfService;
        public string acctBhCode;
        public string acctBhFileDate;
        public string acctBhDischarge;
        public string acctBhProof;
        public string acctBhDismiss;

        public ReturnAccount(string inStatus, string inCreditorGroup, string inDebtorNumber, string inAcctNumber, string inGuarantorName, string inPatientName, string inReturnDescription, string inReturnAmount, string inAmountListed, string inReturnDate, string inLastPayDate, string inLastPayAmount, string inCreditorNumber, string inDateOfService, string inBCode, string inBFileDate, string inBDischarge, string inBProof, string inBDismiss)
        {
            this.acctStatus = inStatus.Trim().ToUpper();
            this.acctCreditorGroup = inCreditorGroup.Trim().ToUpper();
            this.acctDebtorNumber = inDebtorNumber.Trim().ToUpper();
            this.acctAcctNumber = inAcctNumber.Trim().ToUpper();
            this.acctGuarantorName = inGuarantorName.Trim().ToUpper();
            this.acctPatientName = inPatientName.Trim().ToUpper();
            this.acctReturnDescription = inReturnDescription.Trim().ToUpper();
            this.acctReturnAmount = FormatStringDecimal(inReturnAmount.Trim().ToUpper());
            this.acctAmountListed = FormatStringDecimal(inAmountListed.Trim().ToUpper());
            this.acctReturnDate = inReturnDate.Trim().ToUpper();
            this.acctLastPayDate = inLastPayDate.Trim().ToUpper();
            this.acctLastPayAmount = FormatStringDecimal(inLastPayAmount.Trim().ToUpper());
            this.acctCreditorNumber = inCreditorNumber.Trim().ToUpper();
            this.acctDateOfService = inDateOfService.Trim().ToUpper();
            this.acctBhCode = inBCode.Trim().ToUpper();
            this.acctBhFileDate = inBFileDate.Trim().ToUpper();
            this.acctBhDischarge = inBDischarge.Trim().ToUpper();
            this.acctBhProof = inBProof.Trim().ToUpper();
            this.acctBhDismiss = inBDismiss.Trim().ToUpper();
        }

        public ReturnAccount(string inStatus, string inCreditorGroup, string inDebtorNumber, string inAcctNumber, string inGuarantorName, string inPatientName, string inReturnDescription, string inReturnAmount, string inAmountListed, string inReturnDate, string inLastPayDate, string inLastPayAmount, string inCreditorNumber, string inDateOfService)
        {
            this.acctStatus = inStatus.Trim().ToUpper();
            this.acctCreditorGroup = inCreditorGroup.Trim().ToUpper();
            this.acctDebtorNumber = inDebtorNumber.Trim().ToUpper();
            this.acctAcctNumber = inAcctNumber.Trim().ToUpper();
            this.acctGuarantorName = inGuarantorName.Trim().ToUpper();
            this.acctPatientName = inPatientName.Trim().ToUpper();
            this.acctReturnDescription = inReturnDescription.Trim().ToUpper();
            this.acctReturnAmount = FormatStringDecimal(inReturnAmount.Trim().ToUpper());
            this.acctAmountListed = FormatStringDecimal(inAmountListed.Trim().ToUpper());
            this.acctReturnDate = inReturnDate.Trim().ToUpper();
            this.acctLastPayDate = inLastPayDate.Trim().ToUpper();
            this.acctLastPayAmount = FormatStringDecimal(inLastPayAmount.Trim().ToUpper());
            this.acctCreditorNumber = inCreditorNumber.Trim().ToUpper();
            this.acctDateOfService = inDateOfService.Trim().ToUpper();
            this.acctBhCode = "";
            this.acctBhFileDate = "";
            this.acctBhDischarge = "";
            this.acctBhProof = "";
            this.acctBhDismiss = "";
        }

        public static string GetReturnStatus(ReturnAccount returnAccObj)
        {
            String status = returnAccObj.acctStatus;
            return status;
        }

        public static string GetReturnCredGroup(ReturnAccount returnAccObj)
        {
            String credGroup = returnAccObj.acctCreditorGroup;
            return credGroup;
        }

        public static string GetReturnCredNumber(ReturnAccount returnAccObj)
        {
            String credNum = returnAccObj.acctCreditorNumber;
            return credNum;
        }

        public static string GetReturnDbtrNumber(ReturnAccount returnAccObj)
        {
            if (returnAccObj.acctDebtorNumber == null || returnAccObj.acctDebtorNumber.ToString() == "")
            {

            }
            String dbtrNum = FormatAccountNumber(returnAccObj.acctDebtorNumber);
            return dbtrNum;
        }

        public static string GetReturnAcctNumber(ReturnAccount returnAccObj)
        {
            if (returnAccObj.acctAcctNumber == null || returnAccObj.acctAcctNumber.ToString() == "")
            {

            }
            String acctNum = FormatAccountNumber(returnAccObj.acctAcctNumber);
            return acctNum;
        }

        public static string GetReturnGuarantor(ReturnAccount returnAccObj)
        {
            String guar = returnAccObj.acctGuarantorName;
            return guar;
        }

        public static string GetReturnPatient(ReturnAccount returnAccObj)
        {
            String patient = returnAccObj.acctPatientName;
            return patient;
        }

        public static string GetReturnDescription(ReturnAccount returnAccObj)
        {
            String description = returnAccObj.acctReturnDescription;
            return description;
        }

        public static string GetReturnAmount(ReturnAccount returnAccObj)
        {
            String amount = returnAccObj.acctReturnAmount;
            return amount;
        }

        public static string GetReturnAmtListed(ReturnAccount returnAccObj)
        {
            String listedAmt = returnAccObj.acctAmountListed;
            return listedAmt;
        }

        public static DateTime GetReturnDate(ReturnAccount returnAccObj)
        {
            DateTime rDate;
            if (returnAccObj != null && (returnAccObj.acctReturnDate) != null && (Convert.ToDateTime(returnAccObj.acctReturnDate)).ToShortDateString() != "" && (Convert.ToDateTime(returnAccObj.acctReturnDate)).ToShortDateString() != "01/01/1900" && (Convert.ToDateTime(returnAccObj.acctReturnDate)).ToShortDateString() != "1/1/1900" && (Convert.ToDateTime(returnAccObj.acctReturnDate)).ToShortDateString() != "1/1/0001" && (Convert.ToDateTime(returnAccObj.acctReturnDate)).ToShortDateString() != "01/01/0001")
            {
                rDate = FormatStringToDate(returnAccObj.acctReturnDate);
            }
            else
            {
                rDate = Convert.ToDateTime("01/01/0001");
            }
            return rDate;
        }

        public static DateTime GetReturnLastPayDate(ReturnAccount returnAccObj)
        {
            DateTime lastPayDt;
            if (returnAccObj != null && (returnAccObj.acctLastPayDate) != null && (Convert.ToDateTime(returnAccObj.acctLastPayDate)).ToShortDateString() != "" && (Convert.ToDateTime(returnAccObj.acctLastPayDate)).ToShortDateString() != "01/01/1900" && (Convert.ToDateTime(returnAccObj.acctLastPayDate)).ToShortDateString() != "1/1/1900" && (Convert.ToDateTime(returnAccObj.acctLastPayDate)).ToShortDateString() != "1/1/0001" && (Convert.ToDateTime(returnAccObj.acctLastPayDate)).ToShortDateString() != "01/01/0001")
            {
                lastPayDt = FormatStringToDate(returnAccObj.acctLastPayDate);
            }
            else
            {
                lastPayDt = Convert.ToDateTime("01/01/0001");
            }
            return lastPayDt;
        }

        public static string GetReturnLastPayAmount(ReturnAccount returnAccObj)
        {
            String lastPayAmt = returnAccObj.acctLastPayAmount;
            return lastPayAmt;
        }

        public static DateTime GetReturnServiceDate(ReturnAccount returnAccObj)
        {
            DateTime serviceDt = FormatStringToDate((string)returnAccObj.acctDateOfService);
            return serviceDt;
        }

        public static string GetReturnBhCode(ReturnAccount returnAccObj)
        {
            String bhCode = "";
            if (returnAccObj != null && returnAccObj.ToString() != "")
            {
                bhCode = returnAccObj.acctBhCode;
                if (bhCode == null || bhCode.ToString() == "")
                {
                    bhCode = "";
                }
            }
            return bhCode;
        }

        public static DateTime GetReturnBhFileDate(ReturnAccount returnAccObj)
        {
            DateTime fileDate;
            if (returnAccObj != null && returnAccObj.acctBhFileDate != null && (string)returnAccObj.acctBhFileDate != "" && (Convert.ToDateTime(returnAccObj.acctBhFileDate)).ToShortDateString() != "" && (Convert.ToDateTime(returnAccObj.acctBhFileDate)).ToShortDateString() != "01/01/1900" && (Convert.ToDateTime(returnAccObj.acctBhFileDate)).ToShortDateString() != "1/1/1900" && (Convert.ToDateTime(returnAccObj.acctBhFileDate)).ToShortDateString() != "1/1/0001" && (Convert.ToDateTime(returnAccObj.acctBhFileDate)).ToShortDateString() != "01/01/0001")
            {
                fileDate = FormatStringToDate((string)returnAccObj.acctBhFileDate);
            }
            else
            {
               fileDate = Convert.ToDateTime("01/01/0001");
            }
            return fileDate;
        }

        public static string GetReturnBhDischarge(ReturnAccount returnAccObj)
        {
            String discharge = "";
            if (returnAccObj != null && returnAccObj.ToString() != "")
            {
                discharge = returnAccObj.acctBhDischarge;
            }
            return discharge;
        }

        public static string GetReturnBhProof(ReturnAccount returnAccObj)
        {
            String proof = "";
            if (returnAccObj != null)
            {
                proof = returnAccObj.acctBhProof;
                if (proof == null || proof.ToString() == "")
                {
                    proof = "";
                }
            }
            return proof;
        }

        public static DateTime GetReturnBhDismiss(ReturnAccount returnAccObj)
        {
            DateTime dismissDt = Convert.ToDateTime("01/01/1900");
            if (returnAccObj != null)
            {
                String dismiss = returnAccObj.acctBhDismiss;
                if (dismiss != null && dismiss.ToString() != "")
                {
                    dismissDt = Convert.ToDateTime(dismiss);
                }
            }
            return dismissDt;
        }

        public static string FormatStringDecimal(string inStringNumber)
        {
            decimal value = decimal.Parse(inStringNumber, NumberStyles.Currency | NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint); // load with - for negative values!

            inStringNumber = value.ToString();

            if (!(inStringNumber.StartsWith("$")))
            {
                inStringNumber = "$" + inStringNumber;
            }

            return inStringNumber;
        }

        public static decimal FormatStringToDecimal(string inStringNumber)
        {
            if (!(inStringNumber.StartsWith("$")))
            {
                inStringNumber = "$" + inStringNumber;
            }

            decimal value = decimal.Parse(inStringNumber, NumberStyles.Currency | NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint); // load with - for negative values!

            return value;
        }

        public static DateTime FormatStringToDate(string dateStr)
        {
            DateTime date;
            if (dateStr != null && dateStr != "" && (Convert.ToDateTime(dateStr)).ToShortDateString() != "01/01/1900" && (Convert.ToDateTime(dateStr)).ToShortDateString() != "1/1/1900" && (Convert.ToDateTime(dateStr)).ToShortDateString() != "1/1/0001" && (Convert.ToDateTime(dateStr)).ToShortDateString() != "01/01/0001") {
                date = Convert.ToDateTime(dateStr);
            }
            else
            {
                date = Convert.ToDateTime("01/01/0001");
            }
            return date;
        }

        public static string FormatDateToString(DateTime dt)
        {
            string newDateStr = dt.ToShortDateString();
            return newDateStr;
        }

        public static string FormatAccountNumber(String acctNum)
        {
            if (acctNum.Length > 0)
            {
                int counter = 0; // represents position of current character in account number string
                foreach (char character in acctNum)
                {
                    if (character >= 'A' && character <= 'Z' && counter > (acctNum.Length / 2 + 1)) // counter for position of character within account number (don't remove leading letters!)
                    {
                        acctNum = acctNum.Substring(0, counter); // shortened by one each time
                    }
                    counter++;
                }

            }
            else
            {
                acctNum = null;
                throw new Exception("Invalid or missing account number!");
            }
            return acctNum;
        }
    }
}
