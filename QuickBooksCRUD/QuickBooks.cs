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
using System.Diagnostics;
using Newtonsoft.Json.Linq;
namespace QuickBooksCRUD
{
    public class QuickBooks
    {
        public void DoInvoiceAdd(Dictionary<string, List<ItemModel>> data)
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

                BuildInvoiceAddRq(requestMsgSet, data);
                //BuildDepositAddRq(requestMsgSet);

                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;
                Stopwatch stopwatch = new Stopwatch();


                stopwatch.Start();
                Console.WriteLine($"Time before add inovice in QuickBooks : {stopwatch.ElapsedMilliseconds} ms");
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                stopwatch.Stop();
                if (responseMsgSet != null)
                {
                    IResponseList responseList = responseMsgSet.ResponseList;
                    if (responseList != null)
                    {
                        for (int i = 0; i < responseList.Count; i++)
                        {
                            IResponse response = responseList.GetAt(i);
                            if (response.StatusCode == 0) 
                            {
                                Console.WriteLine("Invoice added successfully in QuickBooks.");
                            }
                            else
                            {
                                Console.WriteLine($"Error: {response.StatusMessage} (Code: {response.StatusCode})");
                            }
                        }
                    }
                }

                // Display elapsed time
                Console.WriteLine($"Time taken for add item in QuickBooks : {stopwatch.ElapsedMilliseconds} ms");
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
        void BuildInvoiceAddRq(IMsgSetRequest requestMsgSet, Dictionary<string, List<ItemModel>> data)
        {
            Stopwatch stopwatch = new Stopwatch();

         
            stopwatch.Start();
            foreach (var category in data)
            {

                Console.WriteLine(category.Key);
                {
                    int count=category.Value.Count;
                    Console.WriteLine($"NO of item contain duplicate also : {count} ");
                    foreach (var item in category.Value)
                    {
                        //Console.WriteLine(item.Item +  "  " + item.Price);
                        if (item.Price < 0) 
                        {
                            ICreditMemoAdd CreditMemoAddRq = requestMsgSet.AppendCreditMemoAddRq();
                            CreditMemoAddRq.CustomerRef.FullName.SetValue("Aameer");
                            CreditMemoAddRq.Memo.SetValue(item.Invoice);
                            IORCreditMemoLineAdd ORCreditMemoLineAdd1 = CreditMemoAddRq.ORCreditMemoLineAddList.Append();
                            ORCreditMemoLineAdd1.CreditMemoLineAdd.ItemRef.FullName.SetValue(item.Item);
                            ORCreditMemoLineAdd1.CreditMemoLineAdd.Quantity.SetValue(1);
                            ORCreditMemoLineAdd1.CreditMemoLineAdd.ServiceDate.SetValue(DateTime.UtcNow.AddDays(-3));
                            ORCreditMemoLineAdd1.CreditMemoLineAdd.Amount.SetValue(Math.Abs(item.Price)); 
                        }
                        else 
                        {
                            IInvoiceAdd InvoiceAddRq = requestMsgSet.AppendInvoiceAddRq();
                            InvoiceAddRq.CustomerRef.FullName.SetValue("Aameer");

                            InvoiceAddRq.TxnDate.SetValue(DateTime.UtcNow.AddDays(-3));
                            InvoiceAddRq.Memo.SetValue(item.Invoice);

                            IORInvoiceLineAdd ORInvoiceLineAdd1 = InvoiceAddRq.ORInvoiceLineAddList.Append();
                            ORInvoiceLineAdd1.InvoiceLineAdd.ItemRef.FullName.SetValue(item.Item);
                            ORInvoiceLineAdd1.InvoiceLineAdd.Quantity.SetValue(1);
                            ORInvoiceLineAdd1.InvoiceLineAdd.Amount.SetValue(item.Price);
                        }

                        //IInvoiceAdd InvoiceAddRq = requestMsgSet.AppendInvoiceAddRq();
                        //InvoiceAddRq.CustomerRef.FullName.SetValue("Aameer");
                  
                        //InvoiceAddRq.TxnDate.SetValue(DateTime.UtcNow.AddDays(-1));
                        //InvoiceAddRq.Memo.SetValue(item.Invoice);
                        //IORInvoiceLineAdd ORInvoiceLineAdd1 = InvoiceAddRq.ORInvoiceLineAddList.Append();
                        //ORInvoiceLineAdd1.InvoiceLineAdd.ItemRef.FullName.SetValue(item.Item);
                        //ORInvoiceLineAdd1.InvoiceLineAdd.Quantity.SetValue(1);
                        //ORInvoiceLineAdd1.InvoiceLineAdd.Amount.SetValue(item.Price);
                      
                      
                    }

                }
            }
            stopwatch.Stop();

            Console.WriteLine($"Time taken for add item in BuildInvoiceAddRq : {stopwatch.ElapsedMilliseconds} ms");

        }
        void BuildDepositAddRq(IMsgSetRequest requestMsgSet)
        {
            IDepositAdd DepositAddRq = requestMsgSet.AppendDepositAddRq();

            DepositAddRq.DepositToAccountRef.FullName.SetValue("Checking Account");
            DepositAddRq.TxnDate.SetValue(DateTime.Now);

            IDepositLineAdd DepositLineAdd = DepositAddRq.DepositLineAddList.Append();
            DepositLineAdd.ORDepositLineAdd.DepositInfo.AccountRef.ListID.SetValue("80000026-1738573710");
            DepositLineAdd.ORDepositLineAdd.DepositInfo.Amount.SetValue(4030.00);
            DepositLineAdd.ORDepositLineAdd.DepositInfo.Memo.SetValue("Service Revenue");

        }
        public void DoItemAdd(Dictionary<string, List<ItemModel>> data)
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager? sessionManager = null;

            try
            {
                //Create the session Manager object
                sessionManager = new QBSessionManager();

                //Create the message set request object to hold our request
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                BuildItemServiceAddRq(requestMsgSet, data);

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
            finally
            {
                sessionManager?.EndSession();
                sessionManager?.CloseConnection();
            }
        }
        void BuildItemServiceAddRq(IMsgSetRequest requestMsgSet, Dictionary<string, List<ItemModel>> data)
        {
           
            foreach (var category in data)
            {
                if (category.Key == "Hardware")
                {
                    foreach (var item in category.Value)
                    {
                        Console.WriteLine(item.Item + item.Price);

                        IItemServiceAdd ItemServiceAddRq = requestMsgSet.AppendItemServiceAddRq();

                        ItemServiceAddRq.Name.SetValue(item.Item);
                        ItemServiceAddRq.IsActive.SetValue(true);

                        ItemServiceAddRq.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.SetValue(item.Price);
                        //ItemServiceAddRq.ClassRef.FullName.SetValue("Revenue:Residential revenue:Internet Services:Wireless Services:Cal.net:Fiber Service");

                        //ItemServiceAddRq.ParentRef.FullName.SetValue("Fiber Service");
                        ItemServiceAddRq.ORSalesPurchase.SalesOrPurchase.AccountRef.ListID.SetValue("80000033-1738573943");
                        //ItemServiceAddRq.ORSalesPurchase.SalesOrPurchase.AccountRef.FullName.SetValue("Revenue:Residential revenue:Internet Services:Wireless Services:Cal.net");

                    }

                }
              
            }
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
                        string? list = null;
                        string categoryListID = Convert.ToString(itemList.GetAt(i).ItemServiceRet.FullName.GetValue());
                        var item = itemList.GetAt(i);
                        if (item != null && item.ItemServiceRet != null && item.ItemServiceRet.ClassRef != null)
                        {
                            list = Convert.ToString(item.ItemServiceRet.ClassRef.ListID.GetValue());
                        }
                        else
                        {
                            Console.WriteLine("One of the objects in the chain is null.");
                        }

                        Console.WriteLine($"{name} | List ID: {listID}  | Type:  {type} id:{categoryListID}");
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
                    string phone = companyInfo.Phone != null ? companyInfo.Phone.GetValue() : "N/A";
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
        public void GetClasses2()
        {
            QBSessionManager sessionManager = new QBSessionManager();

            try
            {
                // Step 1: Open QuickBooks Session
                sessionManager.OpenConnection("", "QuickBooks Class Fetcher");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                IMsgSetRequest requestSet = sessionManager.CreateMsgSetRequest("US", 16, 0); // QBSDK 16.0
                requestSet.Attributes.OnError = ENRqOnError.roeContinue;

                IClassQuery classQuery = requestSet.AppendClassQueryRq();  // Create Class Query Request

                IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);

                //IResponse response = responseSet.ResponseList.GetAt(0);
                IResponseList responseList = responseSet.ResponseList;
                for (int i = 0; i < responseList.Count; i++)
                {
                    IResponse response = responseList.GetAt(i);
                    //check the status code of the response, 0=ok, >0 is warning
                    if (response.StatusCode >= 0)
                    {
                        //the request-specific response is in the details, make sure we have some
                        if (response.Detail != null)
                        {
                            //make sure the response is the type we're expecting
                            ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                            if (responseType == ENResponseType.rtClassQueryRs)
                            {
                                //upcast to more specific type here, this is safe because we checked with response.Type check above
                                IClassRetList ClassRet = (IClassRetList)response.Detail;
                                WalkClassRet(ClassRet);
                            }
                        }
                    }
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

        void WalkClassRet2(IClassRetList classList)
        {
            for (int i = 0; i < classList.Count; i++)
            {
                string? classID = (string)classList.GetAt(i).ListID.GetValue();
                string className = (string)classList.GetAt(i).Name.GetValue();

                Console.WriteLine($"{className} | Class ID: {classID}");
            }
        }
        public void GetCategory()
        {
            QBSessionManager sessionManager = new QBSessionManager();

            try
            {
                // Step 1: Open QuickBooks Session
                sessionManager.OpenConnection("", "QuickBooks Category Fetcher");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                // Step 2: Create Request
                IMsgSetRequest requestSet = sessionManager.CreateMsgSetRequest("US", 16, 0); // QBSDK 16.0
                requestSet.Attributes.OnError = ENRqOnError.roeContinue;

                // Step 3: Append Item Query Request
                IItemQuery itemQuery = requestSet.AppendItemQueryRq();

                // Step 4: Send Request to QuickBooks
                IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);

                // Step 5: Process Response
                HashSet<string> categorySet = new HashSet<string>();

                if (responseSet.ResponseList.Count > 0)
                {
                    IResponse response = responseSet.ResponseList.GetAt(0);
                    if (response.StatusCode == 0 && response.Detail != null)
                    {
                        IORItemRetList itemList = (IORItemRetList)response.Detail;

                        Console.WriteLine("Categories in QuickBooks:");
                        for (int i = 0; i < itemList.Count; i++)
                        {
                            IORItemRet item = itemList.GetAt(i);

                            // Extract category details (Parent Items)
                            if (item.ItemServiceRet != null && item.ItemServiceRet.ListID != null)
                            {
                                string categoryName = item.ItemServiceRet.FullName.GetValue();
                                string categoryListID = item.ItemServiceRet.ListID.GetValue();
                                categorySet.Add($"{categoryName} | ListID: {categoryListID}");
                            }
                            else if (item.ItemInventoryRet != null && item.ItemInventoryRet.ListID != null)
                            {
                                string categoryName = item.ItemInventoryRet.FullName.GetValue();
                                string categoryListID = item.ItemInventoryRet.ListID.GetValue();
                                categorySet.Add($"{categoryName} | ListID: {categoryListID}");
                            }
                            else if (item.ItemNonInventoryRet != null && item.ItemNonInventoryRet.ListID != null)
                            {
                                string categoryName = item.ItemNonInventoryRet.FullName.GetValue();
                                string categoryListID = item.ItemNonInventoryRet.ListID.GetValue();
                                categorySet.Add($"{categoryName} | ListID: {categoryListID}");
                            }
                            else if (item.ItemOtherChargeRet != null && item.ItemOtherChargeRet.ListID != null)
                            {
                                string categoryName = item.ItemOtherChargeRet.FullName.GetValue();
                                string categoryListID = item.ItemOtherChargeRet.ListID.GetValue();
                                categorySet.Add($"{categoryName} | ListID: {categoryListID}");
                            }
                        }

                        // Display unique categories
                        foreach (var category in categorySet)
                        {
                            Console.WriteLine(category);
                        }
                    }
                    else
                    {
                        Console.WriteLine("No categories found or error: " + response.StatusMessage);
                    }
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

        public void GetClasses()
        {
            QBSessionManager sessionManager = new QBSessionManager();

            try
            {
                // Step 1: Open QuickBooks Session
                sessionManager.OpenConnection("", "QuickBooks Class Fetcher");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                IMsgSetRequest requestSet = sessionManager.CreateMsgSetRequest("US", 16, 0); // QBSDK 16.0
                requestSet.Attributes.OnError = ENRqOnError.roeContinue;

                IClassQuery classQuery = requestSet.AppendClassQueryRq();  // Create Class Query Request

                IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);

                IResponseList responseList = responseSet.ResponseList;
                for (int i = 0; i < responseList.Count; i++)
                {
                    IResponse response = responseList.GetAt(i);
                    // Check the status code of the response, 0=ok, >0 is warning
                    if (response.StatusCode >= 0)
                    {
                        // Ensure the response type is correct
                        ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                        if (responseType == ENResponseType.rtClassQueryRs)
                        {
                            IClassRetList classRetList = (IClassRetList)response.Detail;

                            if (classRetList != null && classRetList.Count > 0)
                            {
                                WalkClassRet(classRetList);
                            }
                            else
                            {
                                Console.WriteLine("No classes found.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Unexpected response type.");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Error: {response.StatusMessage}");
                    }
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

        void WalkClassRet(IClassRetList classList)
        {
            for (int i = 0; i < classList.Count; i++)
            {
                string? classID = (string)classList.GetAt(i).ListID.GetValue();
                string className = (string)classList.GetAt(i).Name.GetValue();

                Console.WriteLine($"{className} | Class ID: {classID}");
            }
        }


    }
}
