using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Interop.QBFC16;
using System;
using System.Net;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.IO;
namespace QuickBooksCRUD
{
    public class QuickBooks
    {
        public void DoInvoiceAdd()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                //Create the session Manager object
                sessionManager = new QBSessionManager();

                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildInvoiceAddRq(requestMsgSet);

                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                //End the session and close the connection to QuickBooks
                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message, "Error");
                if (sessionBegun)
                {
                    sessionManager?.EndSession();
                }
                if (connectionOpen)
                {
                    sessionManager?.CloseConnection();
                }
            }
        }

        void BuildInvoiceAddRq(IMsgSetRequest requestMsgSet)
        {
            IInvoiceAdd InvoiceAddRq = requestMsgSet.AppendInvoiceAddRq();

            InvoiceAddRq.CustomerRef.FullName.SetValue("Aameer");
            InvoiceAddRq.TxnDate.SetValue(DateTime.Now.AddDays(-1));
            IORInvoiceLineAdd ORInvoiceLineAdd1 = InvoiceAddRq.ORInvoiceLineAddList.Append();
            ORInvoiceLineAdd1.InvoiceLineAdd.ItemRef.FullName.SetValue("BHTowerRent");
            ORInvoiceLineAdd1.InvoiceLineAdd.Quantity.SetValue(1);
            IORInvoiceLineAdd ORInvoiceLineAdd2 = InvoiceAddRq.ORInvoiceLineAddList.Append();
            ORInvoiceLineAdd2.InvoiceLineAdd.ItemRef.FullName.SetValue("SVC#118");
            ORInvoiceLineAdd2.InvoiceLineAdd.Quantity.SetValue(1);
            IORInvoiceLineAdd ORInvoiceLineAdd3 = InvoiceAddRq.ORInvoiceLineAddList.Append();
            ORInvoiceLineAdd3.InvoiceLineAdd.ItemRef.FullName.SetValue("WiFiCredit");
            ORInvoiceLineAdd3.InvoiceLineAdd.Quantity.SetValue(1);
            InvoiceAddRq.Memo.SetValue("Invoice for Aameer");
        }

        public void DoItemAdd()
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager sessionManager = null;

            try
            {
                //Create the session Manager object
                sessionManager = new QBSessionManager();

                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildItemServiceAddRq(requestMsgSet);

                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);

                //End the session and close the connection to QuickBooks
                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message, "Error");
                if (sessionBegun)
                {
                    sessionManager?.EndSession();
                }
                if (connectionOpen)
                {
                    sessionManager?.CloseConnection();
                }
            }
        }



        void BuildItemServiceAddRq(IMsgSetRequest requestMsgSet)
        {
            // Create a service item request
            IItemServiceAdd ItemServiceAddRq = requestMsgSet.AppendItemServiceAddRq();

            ItemServiceAddRq.Name.SetValue("Product B");
            ItemServiceAddRq.IsActive.SetValue(true);

            ItemServiceAddRq.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.SetValue(15.65);

            ItemServiceAddRq.ORSalesPurchase.SalesOrPurchase.AccountRef.ListID.SetValue("80000026-1738573710");

        }

        public void GetAccount()
        {
            QBSessionManager sessionManager = new QBSessionManager();

            try
            {
                // Step 1: Open QuickBooks Session
                sessionManager.OpenConnection("", "QuickBooks Account Fetcher");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                // Step 2: Create Request
                IMsgSetRequest requestSet = sessionManager.CreateMsgSetRequest("US", 16, 0); // QBSDK 16.0
                requestSet.Attributes.OnError = ENRqOnError.roeContinue;

                // Step 3: Append Account Query Request
                IAccountQuery accountQuery = requestSet.AppendAccountQueryRq();

                // Step 4: Send Request to QuickBooks
                IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);

                // Step 5: Process Response
                IResponse response = responseSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    IAccountRetList accountList = (IAccountRetList)response.Detail;

                    Console.WriteLine("Accounts in QuickBooks:");
                    for (int i = 0; i < accountList.Count; i++)
                    {
                        IAccountRet account = accountList.GetAt(i);
                        string name = account.Name.GetValue();
                        string type = account.AccountType.ToString();
                        string listID = account.ListID != null ? account.ListID.GetValue() : "N/A";

                        Console.WriteLine($"{name} | List ID: {listID} ");
                    }
                }
                else
                {
                    Console.WriteLine("No accounts found or error: " + response.StatusMessage);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
            }
            finally
            {
                sessionManager.EndSession();
                sessionManager.CloseConnection();
            }
        }
    }
}
