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



    }
}

