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

        public void DoAccountQuery()
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

                BuildAccountQueryRq(requestMsgSet);

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

        void BuildAccountQueryRq(IMsgSetRequest requestMsgSet)
                {
                        IAccountQuery AccountQueryRq= requestMsgSet.AppendAccountQueryRq();
                        AccountQueryRq.metaData.SetValue("IQBENmetaDataType");
                        string ORAccountListQueryElementType433 = "ListIDList";
                        if (ORAccountListQueryElementType433 == "ListIDList")
                        {
                                //Set field value for ListIDList
                                //May create more than one of these if needed
                                AccountQueryRq.ORAccountListQuery.ListIDList.Add("200000-1011023419");
                        }
                        if (ORAccountListQueryElementType433 == "FullNameList")
                        {
                                //Set field value for FullNameList
                                //May create more than one of these if needed
                                AccountQueryRq.ORAccountListQuery.FullNameList.Add("Protection Plan");
                        }
                        if (ORAccountListQueryElementType433 == "AccountListFilter")
                        {
                                //Set field value for MaxReturned
                                AccountQueryRq.ORAccountListQuery.AccountListFilter.MaxReturned.SetValue(6);
                                //Set field value for ActiveStatus
                                AccountQueryRq.ORAccountListQuery.AccountListFilter.ActiveStatus.SetValue(ENActiveStatus.asActiveOnly);
                                //Set field value for FromModifiedDate
                                AccountQueryRq.ORAccountListQuery.AccountListFilter.FromModifiedDate.SetValue(DateTime.Parse("12/15/2007 12:15:12"),false);
                                //Set field value for ToModifiedDate
                                AccountQueryRq.ORAccountListQuery.AccountListFilter.ToModifiedDate.SetValue(DateTime.Parse("12/15/2007 12:15:12"),false);
                                string ORNameFilterElementType434 = "NameFilter";
                                if (ORNameFilterElementType434 == "NameFilter")
                                {
                                        //Set field value for MatchCriterion
                                        AccountQueryRq.ORAccountListQuery.AccountListFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                                        //Set field value for Name
                                        AccountQueryRq.ORAccountListQuery.AccountListFilter.ORNameFilter.NameFilter.Name.SetValue("ab");
                                }
                                if (ORNameFilterElementType434 == "NameRangeFilter")
                                {
                                        //Set field value for FromName
                                        AccountQueryRq.ORAccountListQuery.AccountListFilter.ORNameFilter.NameRangeFilter.FromName.SetValue("ab");
                                        //Set field value for ToName
                                        AccountQueryRq.ORAccountListQuery.AccountListFilter.ORNameFilter.NameRangeFilter.ToName.SetValue("ab");
                                }
                                //Set field value for AccountTypeList
                                //May create more than one of these if needed
                                //AccountQueryRq.ORAccountListQuery.AccountListFilter.AccountTypeList.Add(ENAccountTypeList.atlAccountsPayable);
                                string ORCurrencyFilterElementType435 = "ListIDList";
                                if (ORCurrencyFilterElementType435 == "ListIDList")
                                {
                                        //Set field value for ListIDList
                                        //May create more than one of these if needed
                                        AccountQueryRq.ORAccountListQuery.AccountListFilter.CurrencyFilter.ORCurrencyFilter.ListIDList.Add("200000-1011023419");
                                }
                                if (ORCurrencyFilterElementType435 == "FullNameList")
                                {
                                        //Set field value for FullNameList
                                        //May create more than one of these if needed
                                        AccountQueryRq.ORAccountListQuery.AccountListFilter.CurrencyFilter.ORCurrencyFilter.FullNameList.Add("ab");
                                }
                        }
                        //Set field value for IncludeRetElementList
                        //May create more than one of these if needed
                        AccountQueryRq.IncludeRetElementList.Add("ab");
                        //Set field value for OwnerIDList
                        //May create more than one of these if needed
                        AccountQueryRq.OwnerIDList.Add(Guid.NewGuid().ToString());
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

                BuildItemInventoryAddRq(requestMsgSet);

                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //Send the request and get the response from QuickBooks
                IMsgSetResponse response = sessionManager.DoRequests(requestMsgSet);
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
        public void DoItemQuery()
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

                BuildItemQueryRq(requestMsgSet);

                //Connect to QuickBooks and begin a session
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;

                //Send the request and get the response from QuickBooks
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                Console.WriteLine(responseMsgSet.ResponseList);
                //End the session and close the connection to QuickBooks
                sessionManager.EndSession();
                sessionBegun = false;
                sessionManager.CloseConnection();
                connectionOpen = false;

            }
            catch (Exception e)
            {
                Console.WriteLine("Error :", e);
                if (sessionBegun)
                {

                    sessionManager.EndSession();
                }
                if (connectionOpen)
                {
                    sessionManager.CloseConnection();
                }
            }
        }
        void BuildItemQueryRq(IMsgSetRequest requestMsgSet)
        {
            //IItemQuery ItemQueryRq = requestMsgSet.AppendItemQueryRq();
            IItemQuery ItemQueryRq = requestMsgSet.AppendItemQueryRq();
            //Set attributes
            //Set field value for metaData
            //ItemQueryRq.metaData.SetValue("IQBENmetaDataType");
            ////Set field value for iterator
            //ItemQueryRq.iterator.SetValue("IQBENiteratorType");
            //Set field value for iteratorID
            ItemQueryRq.iteratorID.SetValue("IQBUUIDType");
            string ORListQueryElementType14039 = "ListIDList";
            if (ORListQueryElementType14039 == "ListIDList")
            {
                //Set field value for ListIDList
                //May create more than one of these if needed
                ItemQueryRq.ORListQuery.ListIDList.Add("200000-1011023419");
            }
            if (ORListQueryElementType14039 == "FullNameList")
            {
                //Set field value for FullNameList
                //May create more than one of these if needed
                ItemQueryRq.ORListQuery.FullNameList.Add("ab");
            }
            if (ORListQueryElementType14039 == "ListFilter")
            {
                //Set field value for MaxReturned
                ItemQueryRq.ORListQuery.ListFilter.MaxReturned.SetValue(6);
                //Set field value for ActiveStatus
                ItemQueryRq.ORListQuery.ListFilter.ActiveStatus.SetValue(ENActiveStatus.asActiveOnly);
                //Set field value for FromModifiedDate
                ItemQueryRq.ORListQuery.ListFilter.FromModifiedDate.SetValue(DateTime.Parse("12/15/2007 12:15:12"), false);
                //Set field value for ToModifiedDate
                ItemQueryRq.ORListQuery.ListFilter.ToModifiedDate.SetValue(DateTime.Parse("12/15/2007 12:15:12"), false);
                string ORNameFilterElementType14040 = "NameFilter";
                if (ORNameFilterElementType14040 == "NameFilter")
                {
                    //Set field value for MatchCriterion
                    ItemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.MatchCriterion.SetValue(ENMatchCriterion.mcStartsWith);
                    //Set field value for Name
                    ItemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameFilter.Name.SetValue("ab");
                }
                if (ORNameFilterElementType14040 == "NameRangeFilter")
                {
                    //Set field value for FromName
                    ItemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameRangeFilter.FromName.SetValue("BHTowerRent");
                    //Set field value for ToName
                    ItemQueryRq.ORListQuery.ListFilter.ORNameFilter.NameRangeFilter.ToName.SetValue("SVC#995");
                }
            }
            //Set field value for IncludeRetElementList
            //May create more than one of these if needed
            ItemQueryRq.IncludeRetElementList.Add("ab");
        }




        void BuildItemInventoryAddRq(IMsgSetRequest requestMsgSet)
        {
            IItemInventoryAdd ItemInventoryAddRq = requestMsgSet.AppendItemInventoryAddRq();
            //Set field value for Name
            ItemInventoryAddRq.Name.SetValue("ab");
            //Set field value for BarCodeValue
            ItemInventoryAddRq.BarCode.BarCodeValue.SetValue("ab");
            //Set field value for AssignEvenIfUsed
            ItemInventoryAddRq.BarCode.AssignEvenIfUsed.SetValue(true);
            //Set field value for AllowOverride
            ItemInventoryAddRq.BarCode.AllowOverride.SetValue(true);
            //Set field value for IsActive
            ItemInventoryAddRq.IsActive.SetValue(true);
            //Set field value for ListID
            ItemInventoryAddRq.ClassRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            ItemInventoryAddRq.ClassRef.FullName.SetValue("ab");
            //Set field value for ListID
            ItemInventoryAddRq.ParentRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            ItemInventoryAddRq.ParentRef.FullName.SetValue("ab");
            //Set field value for ManufacturerPartNumber
            ItemInventoryAddRq.ManufacturerPartNumber.SetValue("ab");
            //Set field value for ListID
            ItemInventoryAddRq.UnitOfMeasureSetRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            ItemInventoryAddRq.UnitOfMeasureSetRef.FullName.SetValue("ab");
            //Set field value for IsTaxIncluded
            ItemInventoryAddRq.IsTaxIncluded.SetValue(true);
            //Set field value for ListID
            ItemInventoryAddRq.SalesTaxCodeRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            ItemInventoryAddRq.SalesTaxCodeRef.FullName.SetValue("ab");
            //Set field value for SalesDesc
            ItemInventoryAddRq.SalesDesc.SetValue("ab");
            //Set field value for SalesPrice
            ItemInventoryAddRq.SalesPrice.SetValue(15.65);
            //Set field value for ListID
            ItemInventoryAddRq.IncomeAccountRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            ItemInventoryAddRq.IncomeAccountRef.FullName.SetValue("ab");
            //Set field value for PurchaseDesc
            ItemInventoryAddRq.PurchaseDesc.SetValue("ab");
            //Set field value for PurchaseCost
            ItemInventoryAddRq.PurchaseCost.SetValue(15.65);
            //Set field value for ListID
            ItemInventoryAddRq.PurchaseTaxCodeRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            ItemInventoryAddRq.PurchaseTaxCodeRef.FullName.SetValue("ab");
            //Set field value for ListID
            ItemInventoryAddRq.COGSAccountRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            ItemInventoryAddRq.COGSAccountRef.FullName.SetValue("ab");
            //Set field value for ListID
            ItemInventoryAddRq.PrefVendorRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            ItemInventoryAddRq.PrefVendorRef.FullName.SetValue("ab");
            //Set field value for ListID
            ItemInventoryAddRq.AssetAccountRef.ListID.SetValue("200000-1011023419");
            //Set field value for FullName
            ItemInventoryAddRq.AssetAccountRef.FullName.SetValue("ab");
            //Set field value for ReorderPoint
            ItemInventoryAddRq.ReorderPoint.SetValue(2);
            //Set field value for Max
            ItemInventoryAddRq.Max.SetValue(2);
            //Set field value for QuantityOnHand
            ItemInventoryAddRq.QuantityOnHand.SetValue(2);
            //Set field value for TotalValue
            ItemInventoryAddRq.TotalValue.SetValue(10.01);
            //Set field value for InventoryDate
            ItemInventoryAddRq.InventoryDate.SetValue(DateTime.Parse("12/15/2007"));
            //Set field value for ExternalGUID
            ItemInventoryAddRq.ExternalGUID.SetValue(Guid.NewGuid().ToString());
            //Set field value for IncludeRetElementList
            //May create more than one of these if needed
            ItemInventoryAddRq.IncludeRetElementList.Add("ab");
        }




    }
}



//void BuildInvoiceAddRq(IMsgSetRequest requestMsgSet)
//{
//    IInvoiceAdd InvoiceAddRq = requestMsgSet.AppendInvoiceAddRq();

//    InvoiceAddRq.CustomerRef.FullName.SetValue("Aameer");
//    InvoiceAddRq.TxnDate.SetValue(DateTime.Now.AddDays(0));
//    IORInvoiceLineAdd ORInvoiceLineAdd1 = InvoiceAddRq.ORInvoiceLineAddList.Append();
//    ORInvoiceLineAdd1.InvoiceLineAdd.ItemRef.FullName.SetValue("BHTowerRent");
//    ORInvoiceLineAdd1.InvoiceLineAdd.Quantity.SetValue(1);
//    IORInvoiceLineAdd ORInvoiceLineAdd2 = InvoiceAddRq.ORInvoiceLineAddList.Append();
//    ORInvoiceLineAdd2.InvoiceLineAdd.ItemRef.FullName.SetValue("SVC1309");
//    ORInvoiceLineAdd2.InvoiceLineAdd.Quantity.SetValue(1);
//    //ORInvoiceLineAdd2.InvoiceLineAdd.OverrideItemAccountRef.FullName.SetValue("Revenue:Residential revenue:Internet Services:Protection Plan");
//    //ORInvoiceLineAdd2.InvoiceLineAdd.ClassRef.FullName.SetValue("Revenue:Residential revenue:Internet Services:Protection Plan:Protection Plan");

//    IORInvoiceLineAdd ORInvoiceLineAdd3 = InvoiceAddRq.ORInvoiceLineAddList.Append();
//    ORInvoiceLineAdd3.InvoiceLineAdd.ItemRef.FullName.SetValue("SVC#1108");
//    ORInvoiceLineAdd3.InvoiceLineAdd.Quantity.SetValue(1);
//    //ORInvoiceLineAdd3.InvoiceLineAdd.OverrideItemAccountRef.FullName.SetValue("Revenue:Residential revenue:Internet Services:Protection Plan");
//    //ORInvoiceLineAdd3.InvoiceLineAdd.ClassRef.FullName.SetValue("Revenue:Residential revenue:Internet Services:Protection Plan:Protection Plan");

//    InvoiceAddRq.Memo.SetValue("Invoice for Aameer");
//}