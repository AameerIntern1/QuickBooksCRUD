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

            ItemServiceAddRq.Name.SetValue("Product 1");
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
                        string name = account.FullName.GetValue();
                        string? type = Convert.ToString(account.AccountType.GetValue());
                        string listID = account.ListID != null ? account.ListID.GetValue() : "N/A";

                        Console.WriteLine($"{name} | List ID: {listID} |  type:{type} ");
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
        public void GetItems()
        {
            QBSessionManager sessionManager = new QBSessionManager();

            try
            {
                // Step 1: Open QuickBooks Session
                sessionManager.OpenConnection("", "QuickBooks Item Fetcher");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                IMsgSetRequest requestSet = sessionManager.CreateMsgSetRequest("US", 16, 0); // QBSDK 16.0
                requestSet.Attributes.OnError = ENRqOnError.roeContinue;

                IItemQuery itemQuery = requestSet.AppendItemQueryRq();

                IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);

                IResponse response = responseSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    IORItemRetList itemList = (IORItemRetList)response.Detail;

                    Console.WriteLine("Items in QuickBooks:");
                    for (int i = 0; i < itemList.Count; i++)
                    {
                        string? listID = (string)itemList.GetAt(i).ItemServiceRet.ListID.GetValue();
                        string name = (string)itemList.GetAt(i).ItemServiceRet.Name.GetValue();
                        string type = Convert.ToString(itemList.GetAt(i).ItemServiceRet.Type.GetValue()); 
                        Console.WriteLine($"{name} | List ID: {listID}  | Type:  {type}");
                    }
                }
                else
                {

                
                    Console.WriteLine("No items found or error: " + response.StatusMessage);
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
        public void GetInvoices()
        {
            QBSessionManager sessionManager = new QBSessionManager();

            try
            {
                // Step 1: Open QuickBooks Session
                sessionManager.OpenConnection("", "QuickBooks Invoice Fetcher");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                // Step 2: Create Request for Invoice Query
                IMsgSetRequest requestSet = sessionManager.CreateMsgSetRequest("US", 16, 0); // QBSDK 16.0
                requestSet.Attributes.OnError = ENRqOnError.roeContinue;

                // Step 3: Append Invoice Query Request
                IInvoiceQuery invoiceQuery = requestSet.AppendInvoiceQueryRq();

                // Step 4: Send Request to QuickBooks
                IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);

                // Step 5: Process Response
                IResponse response = responseSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    IInvoiceRetList invoiceList = (IInvoiceRetList)response.Detail;

                    Console.WriteLine("Invoices in QuickBooks:");
                    for (int i = 0; i < invoiceList.Count; i++)
                    {
                        IInvoiceRet invoice = invoiceList.GetAt(i);
                        string invoiceID = invoice.RefNumber.GetValue();
                        string customerName = invoice.CustomerRef.FullName.GetValue();
                        string balanceAmount = invoice.BalanceRemaining.GetValue().ToString();
                        string totalAmount = invoice.Subtotal.GetValue().ToString();  

                        Console.WriteLine($"Invoice ID: {invoiceID} | Customer: {customerName} | Balance Amount: {balanceAmount} | Paid Amount: {totalAmount}");
                    }
                }
                else
                {
                    Console.WriteLine("No invoices found or error: " + response.StatusMessage);
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
        public void GetCompanyInfo()
        {
            QBSessionManager sessionManager = new QBSessionManager();

            try
            {
                sessionManager.OpenConnection("", "QuickBooks Company Info Fetcher");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                IMsgSetRequest requestSet = sessionManager.CreateMsgSetRequest("US", 16, 0); 
                requestSet.Attributes.OnError = ENRqOnError.roeContinue;

                ICompanyQuery companyQuery = requestSet.AppendCompanyQueryRq();

                IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);

                IResponse response = responseSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    ICompanyRet companyInfo = (ICompanyRet)response.Detail;

                    Console.WriteLine("Company Information in QuickBooks:");
                    string companyName = companyInfo.CompanyName.GetValue();
                    string email = companyInfo.Email != null ? companyInfo.Email.GetValue() : "N/A";
                    string address = companyInfo.Address != null ? companyInfo.Address.Addr1.GetValue() : "N/A";
                    string city = companyInfo.Address != null ? companyInfo.Address.City.GetValue() : "N/A";
                    string phone = companyInfo.Phone !=null ?companyInfo.Phone.GetValue(): "N/A";
                    string fax = companyInfo.Fax != null ? companyInfo.Fax.GetValue() : "N/A";

                    Console.WriteLine($"Company Name: {companyName}");
                    Console.WriteLine($"Email: {email}");
                    Console.WriteLine($"Address: {address}, {city}");
                    Console.WriteLine($"Phone: {phone}");
                    Console.WriteLine($"Fax: {fax}");
                }
                else
                {
                    Console.WriteLine("No company information found or error: " + response.StatusMessage);
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
