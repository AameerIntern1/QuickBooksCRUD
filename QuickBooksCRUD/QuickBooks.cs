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
using QuickBooksCRUD;
using static System.Runtime.InteropServices.JavaScript.JSType;
namespace QuickBooksCRUD
{
    public class QuickBooks
    {

        public List<Journal> GetJournal()
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<Journal> previousPrices = new List<Journal>();
            //QBSessionManager sessionManager = new QBSessionManager();
            //List<PreviousPrice> previousPrices = new List<PreviousPrice>();

            try
            {
                sessionManager.OpenConnection("", "QuickBooks Invoice Fetcher");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                IMsgSetRequest requestSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestSet.Attributes.OnError = ENRqOnError.roeContinue;
                IJournalEntryQuery journalEntryQuery = requestSet.AppendJournalEntryQueryRq();
                //invoiceQuery.ORInvoiceQuery.TxnIDList.Add("652-1738924901");  // Example TxnID
                journalEntryQuery.ORTxnQuery.TxnFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue(DateTime.Parse("02/12/2025"));
                //Set field value for ToTxnDate
                journalEntryQuery.ORTxnQuery.TxnFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.ToTxnDate.SetValue(DateTime.Parse("02/12/2025"));

                journalEntryQuery.IncludeLineItems.SetValue(true);

                IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);
                string date = DateTime.Now.AddDays(-1).ToString("MM/yyyy");

                // Step 5: Process Response
                IResponse response = responseSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    IJournalEntryRetList journalList = (IJournalEntryRetList)response.Detail;
                    Console.WriteLine($"Invoices in QuickBooks {journalList.Count}:");
                    int count = 0;

                    for (int i = 0; i < journalList.Count; i++)
                    {
                        IJournalEntryRet journal = journalList.GetAt(i);
                        Console.WriteLine($"Processing Invoice ID: {journal.RefNumber.GetValue()}");

                        string? memo = journal.Memo != null ? Convert.ToString(journal.Memo.GetValue()) : null;

                        if (journal.ORJournalLineList != null)
                        {
                            count++;
                            DateTime txnDate = Convert.ToDateTime(journal.TxnDate.GetValue());
                            string invoiceID = journal.RefNumber.GetValue();
                            string txnID = journal.TxnID.GetValue();
                            string editID = journal.EditSequence.GetValue();
                            for (int j = 0; j < journal.ORJournalLineList.Count; j++)
                            {
                                IORJournalLine lineItem = journal.ORJournalLineList.GetAt(j);

                                string creditAccount = string.Empty;
                                decimal creditPrice = 0;
                                string creditTxnLineId = string.Empty;

                                string debitAccount = string.Empty;
                                decimal debitPrice = 0;
                                string debitTxnLineId = string.Empty;

                                if (lineItem.JournalCreditLine != null)
                                {
                                    if (lineItem.JournalCreditLine.AccountRef != null)
                                        creditAccount = lineItem.JournalCreditLine.AccountRef.FullName.GetValue();

                                    if (lineItem.JournalCreditLine.Amount != null)
                                        creditPrice = Convert.ToDecimal(lineItem.JournalCreditLine.Amount.GetValue());

                                    creditTxnLineId = lineItem.JournalCreditLine.TxnLineID.GetValue();
                                }

                                if (lineItem.JournalDebitLine != null)
                                {
                                    if (lineItem.JournalDebitLine.AccountRef != null)
                                        debitAccount = lineItem.JournalDebitLine.AccountRef.FullName.GetValue();

                                    if (lineItem.JournalDebitLine.Amount != null)
                                        debitPrice = Convert.ToDecimal(lineItem.JournalDebitLine.Amount.GetValue());

                                    debitTxnLineId = lineItem.JournalDebitLine.TxnLineID.GetValue();
                                }

                                if (!string.IsNullOrEmpty(creditAccount) || !string.IsNullOrEmpty(debitAccount))
                                {
                                    previousPrices.Add(new Journal
                                    {
                                        Id = invoiceID,
                                        TaxId = txnID,
                                        CreditTxnLineId = creditTxnLineId,
                                        DebitTxnLineId = debitTxnLineId,
                                        CreditAccount = creditAccount,
                                        DebitAccount = debitAccount,
                                        DebitPrice = debitPrice,
                                        CreditPrice = creditPrice,
                                        EditSequenceID = editID,
                                        TxnDate = txnDate
                                    });
                                }
                            }



                        }
                        else
                        {
                            Console.WriteLine("No line items found.");
                        }
                    }

                    Console.WriteLine($"After Validation {count}:");
                }
                else
                {
                    Console.WriteLine("No invoices found or error: " + response.StatusMessage);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                sessionManager.EndSession();
                sessionManager.CloseConnection();
            }


            return previousPrices;
        }
        public void DailyJournalAdd(Dictionary<string, decimal> data, List<Journal> previousPrices)
        {
            //, Dictionary<string, decimal> invoiceData
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager? sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
                //DailyBuildInvoiceAddRq(requestMsgSet, data,previousPrices);
                DailyBuildJournaleModRq(requestMsgSet, data, previousPrices);
                //BuildItemServiceAddRq(requestMsgSet, data);
                //BuildDepositAddRq(requestMsgSet);
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                Console.WriteLine($"Time before add inovice in QuickBooks : {stopwatch.ElapsedMilliseconds} ms");
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                stopwatch.Stop();
                int count = 0;
                int invoiceCount = 0;
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
                                invoiceCount++;
                            }
                            else
                            {
                                count++;
                                Console.WriteLine(response.StatusMessage);
                            }
                        }
                    }
                    Console.WriteLine($"{invoiceCount} Invoice added successfully in QuickBooks");

                    Console.WriteLine("No of Invoice not inserted =" + count);
                }
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
        private Journal GetLatestJournalEntry(string txnId)
        {
            try
            {
                QBSessionManager sessionManager = new QBSessionManager();
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;

                IJournalEntryQuery journalQuery = requestMsgSet.AppendJournalEntryQueryRq();
                journalQuery.TxnIDList.Add(txnId);

                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                IResponse response = responseMsgSet.ResponseList.GetAt(0);

                if (response.StatusCode != 0) return null; // Error, no journal found

                IJournalEntryRet journalEntry = (IJournalEntryRet)response.Detail;

                return new Journal
                {
                    TaxId = journalEntry.TxnID.GetValue(),
                    EditSequenceID = journalEntry.EditSequence.GetValue(),
                    TxnDate = journalEntry.TxnDate.GetValue()
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching latest journal entry: {ex.Message}");
                return null;
            }
        }
        void DailyBuildJournaleModRq(IMsgSetRequest requestMsgSet, Dictionary<string, decimal> data, List<Journal> previousPrices)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            foreach (var mod in previousPrices)
            {
                if (mod == null) continue; // Ensure mod is not null

                IJournalEntryMod journalModRq = requestMsgSet.AppendJournalEntryModRq();
                journalModRq.TxnID.SetValue(mod.TaxId);
                journalModRq.TxnDate.SetValue(Convert.ToDateTime(mod.TxnDate));
                journalModRq.EditSequence.SetValue(mod.EditSequenceID);

                decimal earnedRevChange = data.ContainsKey("EarnedRevenue") ? data["EarnedRevenue"] : 0;
                decimal deferredRevChange = data.ContainsKey("DeferredRevenue") ? data["DeferredRevenue"] : 0;
                decimal arChange = data.ContainsKey("AmountReceivable") ? data["AmountReceivable"] : 0;
                decimal cashChange = data.ContainsKey("Cash") ? data["Cash"] : 0;

                // Earned Revenue (Credit if positive, Debit if negative)
                if (earnedRevChange != 0)
                {
                    IJournalLineMod earnedRevLine = journalModRq.JournalLineModList.Append();
                    earnedRevLine.AccountRef.FullName.SetValue("Earned Revenue");
                    earnedRevLine.Amount.SetValue(Convert.ToDouble(Math.Abs(earnedRevChange)));
                    earnedRevLine.JournalLineType.SetValue(earnedRevChange > 0 ? ENJournalLineType.jltCredit : ENJournalLineType.jltDebit);
                }

                // Deferred Revenue (Credit if positive, Debit if negative)
                if (deferredRevChange != 0)
                {
                    IJournalLineMod deferredRevLine = journalModRq.JournalLineModList.Append();
                    deferredRevLine.AccountRef.FullName.SetValue("Deferred Revenue");
                    deferredRevLine.Amount.SetValue(Convert.ToDouble(Math.Abs(deferredRevChange)));
                    deferredRevLine.JournalLineType.SetValue(deferredRevChange > 0 ? ENJournalLineType.jltCredit : ENJournalLineType.jltDebit);
                }

                // Accounts Receivable (Debit if positive, Credit if negative)
                if (arChange != 0)
                {
                    IJournalLineMod arLine = journalModRq.JournalLineModList.Append();
                    arLine.AccountRef.FullName.SetValue("Accounts Receivable");
                    arLine.Amount.SetValue(Convert.ToDouble(Math.Abs(arChange)));
                    arLine.JournalLineType.SetValue(arChange > 0 ? ENJournalLineType.jltDebit : ENJournalLineType.jltCredit);
                }

                // Cash (Debit if positive, Credit if negative)
                if (cashChange != 0)
                {
                    IJournalLineMod cashLine = journalModRq.JournalLineModList.Append();
                    cashLine.AccountRef.FullName.SetValue("Cash");
                    cashLine.Amount.SetValue(Convert.ToDouble(Math.Abs(cashChange)));
                    cashLine.JournalLineType.SetValue(cashChange > 0 ? ENJournalLineType.jltDebit : ENJournalLineType.jltCredit);
                }
            }

            stopwatch.Stop();
            Console.WriteLine($"Time taken for modifying journal entries: {stopwatch.ElapsedMilliseconds} ms");
        }
        void DailyBuildJournaleModRq2(IMsgSetRequest requestMsgSet, Dictionary<string, decimal> data, List<Journal> previousPrices)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            double newAmount = 150.00;

            foreach (var mod in previousPrices)
            {
                if (mod == null) continue; // Ensure mod is not null

                IJournalEntryMod journalModRq = requestMsgSet.AppendJournalEntryModRq();

                // Setting the TxnID (Transaction ID) of the invoice to modify
                journalModRq.TxnID.SetValue(mod.TaxId);
                journalModRq.TxnDate.SetValue(Convert.ToDateTime(mod.TxnDate));
                journalModRq.EditSequence.SetValue(mod.EditSequenceID);

                double creditAmount = mod.CreditPrice.HasValue ? Convert.ToDouble(mod.CreditPrice) : 0;
                double debitAmount = mod.DebitPrice.HasValue ? Convert.ToDouble(mod.DebitPrice) : 0;
                double adjustedAmount = Math.Abs(creditAmount - newAmount);

                if (mod.CreditAccount == "Accounts Receivable" || mod.CreditAccount == "Checking")
                {
                    IJournalLineMod debitLineMod = journalModRq.JournalLineModList.Append();
                    IJournalLineMod creditLineMod = journalModRq.JournalLineModList.Append();

                    debitLineMod.TxnLineID.SetValue(mod.CreditTxnLineId != null ? mod.CreditTxnLineId : mod.DebitTxnLineId);
                    debitLineMod.AccountRef.FullName.SetValue(mod.CreditAccount != null ? mod.CreditAccount : mod.DebitAccount);
                    debitLineMod.Amount.SetValue(adjustedAmount);
                    debitLineMod.JournalLineType.SetValue(ENJournalLineType.jltDebit);

                    creditLineMod.TxnLineID.SetValue(mod.DebitTxnLineId != null ? mod.DebitTxnLineId : mod.CreditTxnLineId);
                    creditLineMod.AccountRef.FullName.SetValue(mod.DebitAccount != null ? mod.DebitAccount : mod.CreditAccount);
                    creditLineMod.Amount.SetValue(adjustedAmount);
                    creditLineMod.JournalLineType.SetValue(ENJournalLineType.jltCredit);
                }
                else if (mod.DebitAccount == "Accounts Receivable" || mod.DebitAccount == "Checking")
                {
                    IJournalLineMod debitLineMod = journalModRq.JournalLineModList.Append();
                    IJournalLineMod creditLineMod = journalModRq.JournalLineModList.Append();

                    creditLineMod.TxnLineID.SetValue(mod.DebitTxnLineId != null ? mod.DebitTxnLineId : mod.CreditTxnLineId);
                    creditLineMod.AccountRef.FullName.SetValue(mod.DebitAccount != null ? mod.DebitAccount : mod.CreditAccount);
                    creditLineMod.Amount.SetValue(adjustedAmount);
                    creditLineMod.JournalLineType.SetValue(ENJournalLineType.jltCredit);

                    debitLineMod.TxnLineID.SetValue(mod.CreditTxnLineId != null ? mod.CreditTxnLineId : mod.DebitTxnLineId);
                    debitLineMod.AccountRef.FullName.SetValue(mod.CreditAccount != null ? mod.CreditAccount : mod.DebitAccount);
                    debitLineMod.Amount.SetValue(adjustedAmount);
                    debitLineMod.JournalLineType.SetValue(ENJournalLineType.jltDebit);
                }
                else
                {
                    if (!string.IsNullOrEmpty(mod.CreditAccount) && mod.CreditPrice.HasValue)
                    {
                        IJournalLineMod creditLineMod = journalModRq.JournalLineModList.Append();
                        creditLineMod.TxnLineID.SetValue(mod.CreditTxnLineId);
                        creditLineMod.AccountRef.FullName.SetValue(mod.CreditAccount);
                        creditLineMod.Amount.SetValue(Math.Abs(creditAmount));
                        creditLineMod.JournalLineType.SetValue(ENJournalLineType.jltCredit);
                    }

                    if (!string.IsNullOrEmpty(mod.DebitAccount) && mod.DebitPrice.HasValue)
                    {
                        IJournalLineMod debitLineMod = journalModRq.JournalLineModList.Append();
                        debitLineMod.TxnLineID.SetValue(mod.DebitTxnLineId);
                        debitLineMod.AccountRef.FullName.SetValue(mod.DebitAccount);
                        debitLineMod.Amount.SetValue(Math.Abs(debitAmount));
                        debitLineMod.JournalLineType.SetValue(ENJournalLineType.jltDebit);
                    }
                }
            }

            stopwatch.Stop();
            Console.WriteLine($"Time taken for modifying journal entries: {stopwatch.ElapsedMilliseconds} ms");
        }

        void DailyBuildJournaleModRq3(IMsgSetRequest requestMsgSet, Dictionary<string, decimal> data, List<Journal> previousPrices)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            double newAmount = 150.00;
            foreach (var mod in previousPrices)
            {
                if (mod == null) continue; // Ensure mod is not null

                IJournalEntryMod journalModRq = requestMsgSet.AppendJournalEntryModRq();

                // Setting the TxnID (Transaction ID) of the invoice to modify
                journalModRq.TxnID.SetValue(mod.TaxId);
                journalModRq.TxnDate.SetValue(Convert.ToDateTime(mod.TxnDate));
                journalModRq.EditSequence.SetValue(mod.EditSequenceID);

                if (mod.CreditAccount == "Accounts Receivable" || mod.CreditAccount == "Checking")
                {
                    double amount = Convert.ToDouble(mod.CreditPrice) - newAmount;
                    IJournalLineMod creditLineMod = journalModRq.JournalLineModList.Append();
                    IJournalLineMod debitLineMod = journalModRq.JournalLineModList.Append();
                    if (amount > 0)
                    {

                        debitLineMod.TxnLineID.SetValue(mod.CreditTxnLineId != null ? mod.CreditTxnLineId : mod.DebitTxnLineId);
                        debitLineMod.AccountRef.FullName.SetValue(mod.CreditAccount != null ? mod.CreditAccount : mod.DebitAccount);
                        debitLineMod.Amount.SetValue(Math.Abs(amount));
                        debitLineMod.JournalLineType.SetValue(ENJournalLineType.jltDebit);
                        creditLineMod.Amount.SetValue(0);
                        creditLineMod.JournalLineType.SetValue(ENJournalLineType.jltCredit);
                    }
                    else
                    {

                        creditLineMod.TxnLineID.SetValue(mod.DebitTxnLineId != null ? mod.DebitTxnLineId : mod.CreditTxnLineId);
                        creditLineMod.AccountRef.FullName.SetValue(mod.DebitAccount != null ? mod.DebitAccount : mod.CreditAccount);
                        creditLineMod.Amount.SetValue(Math.Abs(amount));
                        creditLineMod.JournalLineType.SetValue(ENJournalLineType.jltCredit);
                        debitLineMod.Amount.SetValue(0);
                        debitLineMod.JournalLineType.SetValue(ENJournalLineType.jltDebit);
                    }
                }
               else if (mod.DebitAccount == "Accounts Receivable" || mod.DebitAccount == "Checking")
                {
                    double amount = Convert.ToDouble(mod.CreditPrice) - newAmount;
                    IJournalLineMod creditLineMod = journalModRq.JournalLineModList.Append();
                    IJournalLineMod debitLineMod = journalModRq.JournalLineModList.Append();
                    if (amount > 0)
                    {

                        debitLineMod.TxnLineID.SetValue(mod.CreditTxnLineId != null ? mod.CreditTxnLineId : mod.DebitTxnLineId);
                        debitLineMod.AccountRef.FullName.SetValue(mod.CreditAccount != null ? mod.CreditAccount : mod.DebitAccount);
                        debitLineMod.Amount.SetValue(Math.Abs(amount));
                        debitLineMod.JournalLineType.SetValue(ENJournalLineType.jltDebit);
                        creditLineMod.Amount.SetValue(0);
                        creditLineMod.JournalLineType.SetValue(ENJournalLineType.jltCredit);
                    }
                    else
                    {

                        creditLineMod.TxnLineID.SetValue(mod.DebitTxnLineId != null ? mod.DebitTxnLineId : mod.CreditTxnLineId);
                        creditLineMod.AccountRef.FullName.SetValue(mod.DebitAccount != null ? mod.DebitAccount : mod.CreditAccount);
                        creditLineMod.Amount.SetValue(Math.Abs(amount));
                        creditLineMod.JournalLineType.SetValue(ENJournalLineType.jltCredit);
                        debitLineMod.Amount.SetValue(0);
                        debitLineMod.JournalLineType.SetValue(ENJournalLineType.jltDebit);
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(mod.CreditAccount) && mod.CreditPrice.HasValue)
                    {
                        IJournalLineMod creditLineMod = journalModRq.JournalLineModList.Append();
                        creditLineMod.TxnLineID.SetValue(mod.CreditTxnLineId);
                        creditLineMod.AccountRef.FullName.SetValue(mod.CreditAccount);
                        creditLineMod.Amount.SetValue(Math.Abs(Convert.ToDouble(mod.CreditPrice)));
                        creditLineMod.JournalLineType.SetValue(ENJournalLineType.jltCredit);
                    }

                    if (!string.IsNullOrEmpty(mod.DebitAccount) && mod.DebitPrice.HasValue)
                    {
                        IJournalLineMod debitLineMod = journalModRq.JournalLineModList.Append();
                        debitLineMod.TxnLineID.SetValue(mod.DebitTxnLineId);
                        debitLineMod.AccountRef.FullName.SetValue(mod.DebitAccount);
                        debitLineMod.Amount.SetValue(Math.Abs(Convert.ToDouble(mod.DebitPrice)));
                        debitLineMod.JournalLineType.SetValue(ENJournalLineType.jltDebit);
                    }

                }
             
            }

            stopwatch.Stop();
            Console.WriteLine($"Time taken for modifying journal entries: {stopwatch.ElapsedMilliseconds} ms");
        }

        void DailyBuildJournaleModRq1(IMsgSetRequest requestMsgSet, Dictionary<string, decimal> data, List<Journal> previousPrices)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            string date = DateTime.Now.AddDays(-1).ToString("MM/yyyy");

            foreach (var mod in previousPrices)
            {
                if (mod == null) continue; // Ensure mod is not null

                IJournalEntryMod journalModRq = requestMsgSet.AppendJournalEntryModRq();

                // Setting the TxnID (Transaction ID) of the invoice to modify
                journalModRq.TxnID.SetValue(mod.TaxId);
                journalModRq.TxnDate.SetValue(Convert.ToDateTime(mod.TxnDate));
                journalModRq.EditSequence.SetValue(mod.EditSequenceID);

                // Ensure that debit and credit lines are added correctly
                if (!string.IsNullOrEmpty(mod.CreditAccount) && mod.CreditPrice.HasValue)
                {
                    IJournalLineMod creditLineMod = journalModRq.JournalLineModList.Append();
                    creditLineMod.TxnLineID.SetValue(mod.CreditTxnLineId);
                    creditLineMod.AccountRef.FullName.SetValue(mod.CreditAccount);
                    creditLineMod.Amount.SetValue(-Math.Abs(Convert.ToDouble(mod.CreditPrice))); // Ensure it's negative for credit
                }

                if (!string.IsNullOrEmpty(mod.DebitAccount) && mod.DebitPrice.HasValue)
                {
                    IJournalLineMod debitLineMod = journalModRq.JournalLineModList.Append();
                    debitLineMod.TxnLineID.SetValue(mod.DebitTxnLineId);
                    debitLineMod.AccountRef.FullName.SetValue(mod.DebitAccount);
                    debitLineMod.Amount.SetValue(Math.Abs(Convert.ToDouble(mod.DebitPrice))); // Ensure it's positive for debit
                }
            }


            //foreach (var mod in previousPrices)
            //{
            //    //double amount = Convert.ToDouble(mod?.OldPrice + mod?.CreditPrice);

            //    IJournalEntryMod journalModRq = requestMsgSet.AppendJournalEntryModRq();
            //    // Setting the TxnID (Transaction ID) of the invoice to modify
            //    journalModRq.TxnID.SetValue(mod?.TaxId);
            //    //InvoiceModRq.RefNumber.SetValue(mod?.TaxId);
            //    journalModRq.TxnDate.SetValue(Convert.ToDateTime(mod?.TxnDate));
            //    journalModRq.EditSequence.SetValue(mod?.EditSequenceID);
            //    //InvoiceModRq.Memo.SetValue($"{date}-{mod?.Item}");

            //    // Modifying an existing line item or adding a new one
            //    IJournalLineMod creditLineMod = journalModRq.JournalLineModList.Append();
            //    creditLineMod.TxnLineID.SetValue(mod?.CreditTxnLineId);
            //    creditLineMod.AccountRef.FullName.SetValue(mod?.CreditAccount);
            //    creditLineMod.Amount.SetValue(Convert.ToDouble(mod?.CreditPrice));

            //    IJournalLineMod debitLineMod = journalModRq.JournalLineModList.Append();

            //    debitLineMod.TxnLineID.SetValue(mod?.DebitTxnLineId);
            //    debitLineMod.AccountRef.FullName.SetValue(mod?.DebitAccount);
            //    debitLineMod.Amount.SetValue(Convert.ToDouble(mod?.DebitPrice));


            //    //ORInvoiceLineMod1.InvoiceLineMod.OverrideUOMSetRef.FullName.SetValue(mod?.Item);

            //    //ORInvoiceLineMod1.InvoiceLineMod.Amount.SetValue(amount);
            //    //Console.WriteLine($"Txn id={mod?.Id} EDitid={mod?.EditSequenceID} item={mod?.Item} refnumber={mod?.TaxId} previous amount{mod?.OldPrice} new amount{mod?.NewPrice} total = {amount}");
            //}


            stopwatch.Stop();
            Console.WriteLine($"Time taken for modifying invoice in InvoiceMod2_1: {stopwatch.ElapsedMilliseconds} ms");
        }

        public List<PreviousPrice> GetInvoices1(Dictionary<string, decimal> data)
        {
            QBSessionManager sessionManager = new QBSessionManager();
            List<PreviousPrice> previousPrices = new List<PreviousPrice>();
            //QBSessionManager sessionManager = new QBSessionManager();
            //List<PreviousPrice> previousPrices = new List<PreviousPrice>();

            try
            {
                sessionManager.OpenConnection("", "QuickBooks Invoice Fetcher");
                sessionManager.BeginSession("", ENOpenMode.omDontCare);

                IMsgSetRequest requestSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestSet.Attributes.OnError = ENRqOnError.roeContinue;
                IInvoiceQuery invoiceQuery = requestSet.AppendInvoiceQueryRq();
                //invoiceQuery.ORInvoiceQuery.TxnIDList.Add("652-1738924901");  // Example TxnID

                invoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue(DateTime.Parse("02/07/2025"));
                invoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.ToTxnDate.SetValue(DateTime.Parse("02/07/2025"));
                invoiceQuery.ORInvoiceQuery.InvoiceFilter.EntityFilter.OREntityFilter.FullNameList.Add("Test");

                invoiceQuery.IncludeLineItems.SetValue(true);

                IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);
                string date = DateTime.Now.AddDays(-1).ToString("MM/yyyy");

                // Step 5: Process Response
                IResponse response = responseSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    IInvoiceRetList invoiceList = (IInvoiceRetList)response.Detail;
                    Console.WriteLine($"Invoices in QuickBooks {invoiceList.Count}:");
                    int count = 0;

                    for (int i = 0; i < invoiceList.Count; i++)
                    {
                        IInvoiceRet invoice = invoiceList.GetAt(i);
                        Console.WriteLine($"Processing Invoice ID: {invoice.RefNumber.GetValue()}");

                        string? memo = invoice.Memo != null ? Convert.ToString(invoice.Memo.GetValue()) : null;

                        if (invoice.ORInvoiceLineRetList != null)
                        {
                            foreach (var category in data)
                            {
                                if (memo != null && memo == $"{date}-{category.Key}")
                                {
                                    count++;
                                    decimal totalAmount = Convert.ToDecimal(invoice.Subtotal.GetValue());
                                    DateTime txnDate = Convert.ToDateTime(invoice.TxnDate.GetValue());
                                    string invoiceID = invoice.RefNumber.GetValue();
                                    string txnID = invoice.TxnID.GetValue();
                                    string editID = invoice.EditSequence.GetValue();
                                    for (int j = 0; j < invoice.ORInvoiceLineRetList.Count; j++)
                                    {
                                        IORInvoiceLineRet lineItem = invoice.ORInvoiceLineRetList.GetAt(j);
                                        if (lineItem.InvoiceLineRet != null && lineItem.InvoiceLineRet.ItemRef != null)
                                        {
                                            string itemName = lineItem.InvoiceLineRet.ItemRef.FullName.GetValue();
                                            decimal itemPrice = lineItem.InvoiceLineRet.Amount != null ? Convert.ToDecimal(lineItem.InvoiceLineRet.Amount.GetValue()) : 0;
                                            string txnLineId = lineItem.InvoiceLineRet.TxnLineID.GetValue();
                                            Console.WriteLine($"Item Name: {itemName}, Price: {itemPrice}");

                                            previousPrices.Add(new PreviousPrice
                                            {
                                                Id = invoiceID,
                                                TaxId = txnID,
                                                TxnLineId = txnLineId,
                                                Item = itemName,
                                                OldPrice = itemPrice,
                                                NewPrice = category.Value,
                                                EditSequenceID = editID,
                                                TxnDate = txnDate
                                            });
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("No line items found.");
                        }
                    }

                    Console.WriteLine($"After Validation {count}:");
                }
                else
                {
                    Console.WriteLine("No invoices found or error: " + response.StatusMessage);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                sessionManager.EndSession();
                sessionManager.CloseConnection();
            }


            return previousPrices;
        }

        public void DailyInvoiceAdd(Dictionary<string, decimal> data, List<PreviousPrice> previousPrices)
        {
            //, Dictionary<string, decimal> invoiceData
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager? sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
                //DailyBuildInvoiceAddRq(requestMsgSet, data,previousPrices);
                DailyBuildInvoiceModRq(requestMsgSet, data, previousPrices);
                //BuildItemServiceAddRq(requestMsgSet, data);
                //BuildDepositAddRq(requestMsgSet);
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                Console.WriteLine($"Time before add inovice in QuickBooks : {stopwatch.ElapsedMilliseconds} ms");
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                stopwatch.Stop();
                int count = 0;
                int invoiceCount = 0;
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
                                invoiceCount++;
                            }
                            else
                            {
                                count++;
                                Console.WriteLine(response.StatusCode);
                            }
                        }
                    }
                    Console.WriteLine($"{invoiceCount} Invoice added successfully in QuickBooks");

                    Console.WriteLine("No of Invoice not inserted =" + count);
                }
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
        void DailyBuildInvoiceModRq(IMsgSetRequest requestMsgSet, Dictionary<string, decimal> data, List<PreviousPrice> previousPrices)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            string date = DateTime.Now.AddDays(-1).ToString("MM/yyyy");

            foreach (var mod in previousPrices)
            {
                Console.WriteLine($"Modifying Invoice - TxnID: {mod.TaxId}, EditSequence: {mod.EditSequenceID}, Item: {mod.Item}, Old Price: {mod.OldPrice}, New Price: {mod.NewPrice}");
            }

            foreach (var mod in previousPrices)
            {
                double amount = Convert.ToDouble(mod?.OldPrice + mod?.NewPrice);

                IInvoiceMod InvoiceModRq = requestMsgSet.AppendInvoiceModRq();
                InvoiceModRq.CustomerRef.FullName.SetValue("Test");
                // Setting the TxnID (Transaction ID) of the invoice to modify
                InvoiceModRq.TxnID.SetValue(mod?.TaxId);
                //InvoiceModRq.RefNumber.SetValue(mod?.TaxId);
                InvoiceModRq.TxnDate.SetValue(Convert.ToDateTime(mod?.TxnDate));
                InvoiceModRq.EditSequence.SetValue(mod?.EditSequenceID);
                InvoiceModRq.Memo.SetValue($"{date}-{mod?.Item}");

                // Modifying an existing line item or adding a new one
                IORInvoiceLineMod ORInvoiceLineMod1 = InvoiceModRq.ORInvoiceLineModList.Append();
                ORInvoiceLineMod1.InvoiceLineMod.TxnLineID.SetValue(mod?.TxnLineId);
                ORInvoiceLineMod1.InvoiceLineMod.ItemRef.FullName.SetValue(mod?.Item);
                //ORInvoiceLineMod1.InvoiceLineMod.OverrideUOMSetRef.FullName.SetValue(mod?.Item);

                ORInvoiceLineMod1.InvoiceLineMod.Quantity.SetValue(1);
                ORInvoiceLineMod1.InvoiceLineMod.Amount.SetValue(amount);
                Console.WriteLine($"Txn id={mod?.Id} EDitid={mod?.EditSequenceID} item={mod?.Item} refnumber={mod?.TaxId} previous amount{mod?.OldPrice} new amount{mod?.NewPrice} total = {amount}");
            }


            stopwatch.Stop();
            Console.WriteLine($"Time taken for modifying invoice in InvoiceMod2_1: {stopwatch.ElapsedMilliseconds} ms");
        }
        void DailyBuildInvoiceAddRq(IMsgSetRequest requestMsgSet, Dictionary<string, decimal> data, List<PreviousPrice> previousPrices)

        {

            string date = DateTime.Now.AddDays(-1).ToString("MM/yyyy");
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            foreach (var category in data)
            {
                Console.WriteLine($"category = {category.Key}  Total = {category.Value}");
                //if (category.Value < 0)
                //{
                //    ICreditMemoAdd CreditMemoAddRq = requestMsgSet.AppendCreditMemoAddRq();
                //    CreditMemoAddRq.CustomerRef.FullName.SetValue("Test1");
                //    IORCreditMemoLineAdd ORCreditMemoLineAdd1 = CreditMemoAddRq.ORCreditMemoLineAddList.Append();
                //    ORCreditMemoLineAdd1.CreditMemoLineAdd.ItemRef.FullName.SetValue(item.Item);
                //    ORCreditMemoLineAdd1.CreditMemoLineAdd.Quantity.SetValue(1);
                //    ORCreditMemoLineAdd1.CreditMemoLineAdd.ServiceDate.SetValue(Convert.ToDateTime(item.Date));
                //    ORCreditMemoLineAdd1.CreditMemoLineAdd.Amount.SetValue(Math.Abs(item.Price));
                //}


                IInvoiceAdd InvoiceAddRq = requestMsgSet.AppendInvoiceAddRq();
                InvoiceAddRq.CustomerRef.FullName.SetValue("Test");
                InvoiceAddRq.Memo.SetValue($"{date}-{category.Key}");
                IORInvoiceLineAdd ORInvoiceLineAdd1 = InvoiceAddRq.ORInvoiceLineAddList.Append();
                ORInvoiceLineAdd1.InvoiceLineAdd.ItemRef.FullName.SetValue(category.Key);
                ORInvoiceLineAdd1.InvoiceLineAdd.Quantity.SetValue(1);
                ORInvoiceLineAdd1.InvoiceLineAdd.Amount.SetValue(Convert.ToDouble(category.Value));

                //    }
                //}
            }
            stopwatch.Stop();
            Console.WriteLine($"Time taken for add Invoice in BuildInvoiceAddRq : {stopwatch.ElapsedMilliseconds} ms");
        }

        public void DoInvoiceAdd(Dictionary<string, List<ItemModel>> data)
        {
            bool sessionBegun = false;
            bool connectionOpen = false;
            QBSessionManager? sessionManager = null;

            try
            {
                sessionManager = new QBSessionManager();
                IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", 16, 0);
                requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
                BuildInvoiceAddRq(requestMsgSet, data);
                //BuildItemServiceAddRq(requestMsgSet, data);
                //BuildDepositAddRq(requestMsgSet);
                sessionManager.OpenConnection("", "Sample Code from OSR");
                connectionOpen = true;
                sessionManager.BeginSession("", ENOpenMode.omDontCare);
                sessionBegun = true;
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                Console.WriteLine($"Time before add inovice in QuickBooks : {stopwatch.ElapsedMilliseconds} ms");
                IMsgSetResponse responseMsgSet = sessionManager.DoRequests(requestMsgSet);
                stopwatch.Stop();
                int count = 0;
                int invoiceCount = 0;
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
                                invoiceCount++;
                            }
                            else
                            {
                                count++;
                            }
                        }
                    }
                    Console.WriteLine($"{invoiceCount} Invoice added successfully in QuickBooks");

                    Console.WriteLine("No of Invoice not inserted =" + count);
                }
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
                int count = category.Value.Count;
                Console.WriteLine($"Total No of item (Duplicate also included): {count} ");
                foreach (var item in category.Value)
                {
                    if (item.Price < 0)
                    {
                        ICreditMemoAdd CreditMemoAddRq = requestMsgSet.AppendCreditMemoAddRq();
                        CreditMemoAddRq.CustomerRef.FullName.SetValue("Test1");
                        CreditMemoAddRq.Memo.SetValue(item.Invoice);
                        IORCreditMemoLineAdd ORCreditMemoLineAdd1 = CreditMemoAddRq.ORCreditMemoLineAddList.Append();
                        ORCreditMemoLineAdd1.CreditMemoLineAdd.ItemRef.FullName.SetValue(item.Item);
                        ORCreditMemoLineAdd1.CreditMemoLineAdd.Quantity.SetValue(1);
                        ORCreditMemoLineAdd1.CreditMemoLineAdd.ServiceDate.SetValue(Convert.ToDateTime(item.Date));
                        ORCreditMemoLineAdd1.CreditMemoLineAdd.Amount.SetValue(Math.Abs(item.Price));
                    }
                    else
                    {
                        IInvoiceAdd InvoiceAddRq = requestMsgSet.AppendInvoiceAddRq();
                        InvoiceAddRq.CustomerRef.FullName.SetValue("Test1");
                        InvoiceAddRq.TxnDate.SetValue((Convert.ToDateTime(item.Date)));
                        InvoiceAddRq.Memo.SetValue(item.Invoice);
                        IORInvoiceLineAdd ORInvoiceLineAdd1 = InvoiceAddRq.ORInvoiceLineAddList.Append();
                        ORInvoiceLineAdd1.InvoiceLineAdd.ItemRef.FullName.SetValue(item.Item);
                        ORInvoiceLineAdd1.InvoiceLineAdd.Quantity.SetValue(1);
                        ORInvoiceLineAdd1.InvoiceLineAdd.Amount.SetValue(item.Price);
                    }
                }
            }
            stopwatch.Stop();
            Console.WriteLine($"Time taken for add Invoice in BuildInvoiceAddRq : {stopwatch.ElapsedMilliseconds} ms");
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
                                Console.WriteLine("Invoice error in QuickBooks.");

                            }
                        }
                    }
                }
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
                int count = category.Value.Count;
                Console.WriteLine($"NO of item contain duplicate also : {count} ");
                Console.WriteLine($"NO of item contain duplicate also : {category.Key} ");
                foreach (var item in category.Value)
                {
                    IItemServiceAdd ItemServiceAddRq = requestMsgSet.AppendItemServiceAddRq();
                    ItemServiceAddRq.Name.SetValue(item.Item);
                    ItemServiceAddRq.IsActive.SetValue(true);
                    ItemServiceAddRq.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.SetValue(item.Price);
                    //ItemServiceAddRq.ORSalesPurchase.SalesOrPurchase.AccountRef.ListID.SetValue("80000033-1738573943");
                    ItemServiceAddRq.ORSalesPurchase.SalesOrPurchase.AccountRef.FullName.SetValue(category.Key);
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
        public void GetInvoices(Dictionary<string, decimal> data)
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
                invoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue(DateTime.Parse("02/07/2025"));
                invoiceQuery.ORInvoiceQuery.InvoiceFilter.EntityFilter.OREntityFilter.FullNameList.Add("Test");
                invoiceQuery.IncludeLineItems.SetValue(true);

                // Step 4: Send Request to QuickBooks
                IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);
                string date = DateTime.Now.AddDays(-1).ToString("MM/yyyy");

                // Step 5: Process Response
                IResponse response = responseSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    IInvoiceRetList invoiceList = (IInvoiceRetList)response.Detail;
                    Console.WriteLine($"Invoices in QuickBooks {invoiceList.Count}:");
                    int count = 0;

                    // Loop through the invoices
                    for (int i = 0; i < invoiceList.Count; i++)
                    {
                        IInvoiceRet invoice = invoiceList.GetAt(i);
                        Console.WriteLine(invoice);
                        string? memo = invoice.Memo != null ? Convert.ToString(invoice.Memo.GetValue()) : null;
                        Console.WriteLine($"Processing Invoice ID: {invoice.RefNumber.GetValue()}");
                        Console.WriteLine($"Processing Invoice ID: {invoice.RefNumber.GetValue()}");
                        var list = invoice.ORInvoiceLineRetList;
                        //Console.WriteLine($"class Invoice ID: {invoice.}");
                        //if (invoice.ORInvoiceLineRetList!=null)
                        //{
                        //    Console.WriteLine($"Line Item Count: {invoice.ORInvoiceLineRetList.Count}");
                        //}
                        //else
                        //{
                        //    Console.WriteLine("ORInvoiceLineRetList is NULL!");
                        //}

                        if (invoice.ORInvoiceLineRetList != null)
                        {
                            for (int j = 0; j < invoice.ORInvoiceLineRetList.Count; j++)
                            {
                                IORInvoiceLineRet lineItem = invoice.ORInvoiceLineRetList.GetAt(j);
                                if (lineItem.InvoiceLineRet != null && lineItem.InvoiceLineRet.ItemRef != null)
                                {
                                    string itemName = lineItem.InvoiceLineRet.ItemRef.FullName.GetValue();
                                    Console.WriteLine($"Item Name: {itemName}");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("item does not exisist ");
                        }
                        foreach (var category in data)
                        {
                            if (memo != null && memo == $"{date}-{category.Key}")
                            {
                                count++;
                                string invoiceID = invoice.RefNumber.GetValue();
                                string customerName = invoice.CustomerRef.FullName.GetValue();
                                string balanceAmount = invoice.BalanceRemaining.GetValue().ToString();
                                decimal totalAmount = Convert.ToDecimal(invoice.Subtotal.GetValue());
                                string txnDate = invoice.TxnDate.GetValue().ToString();

                                //Console.WriteLine($"Invoice ID: {invoiceID} | Customer: {customerName} | Balance Amount: {balanceAmount} | Paid Amount: {totalAmount} | Txn Date: {txnDate}");
                                //            Console.WriteLine($"Item count: {invoice?.ORInvoiceLineRetList?.Count}");

                                //IORInvoiceLineRetList linesItem = invoice.ORInvoiceLineRetList;
                                // Check if invoice has line items

                            }
                        }
                    }

                    Console.WriteLine($"After Validation {count}:");
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


        public void GetInvoices12(Dictionary<string, decimal> data)
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
                //invoiceQuery.ORInvoiceQuery.InvoiceFilter.SetValue("Your Memo Text");
                invoiceQuery.ORInvoiceQuery.InvoiceFilter.ORDateRangeFilter.TxnDateRangeFilter.ORTxnDateRangeFilter.TxnDateFilter.FromTxnDate.SetValue(DateTime.Parse("02/07/2025"));

                invoiceQuery.ORInvoiceQuery.InvoiceFilter.EntityFilter.OREntityFilter.FullNameList.Add("Test1");
                // Step 4: Send Request to QuickBooks
                IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);
                string date = DateTime.Now.AddDays(-1).ToString("MM/yyyy");
                // Step 5: Process Response
                IResponse response = responseSet.ResponseList.GetAt(0);
                if (response.StatusCode == 0 && response.Detail != null)
                {
                    IInvoiceRetList invoiceList = (IInvoiceRetList)response.Detail;
                    Console.WriteLine($"Invoices in QuickBooks {invoiceList.Count}:");
                    int count = 0;
                    // Loop through the invoices
                    for (int i = 0; i < invoiceList.Count; i++)
                    {
                        foreach (var category in data)
                        {
                            IInvoiceRet invoice = invoiceList.GetAt(i);
                            string? memo = invoice.Memo != null ? Convert.ToString(invoice.Memo.GetValue()) : null;
                            if (memo != null)
                            {
                                if (memo == $"{date}-{category.Key}")
                                {
                                    count++;
                                    string invoiceID = invoice.RefNumber.GetValue();
                                    string customerName = invoice.ClassRef.FullName.GetValue();
                                    string balanceAmount = invoice.BalanceRemaining.GetValue().ToString();
                                    string totalAmount = invoice.Subtotal.GetValue().ToString();
                                    string txnDate = invoice.TxnDate.GetValue().ToString();

                                    Console.WriteLine($"Invoice ID: {invoiceID} | Customer: {customerName} | Balance Amount: {balanceAmount} | Paid Amount: {totalAmount} | Txn Date: {txnDate}");

                                    // Loop through the line items to get item names

                                }
                            }
                        }
                    }

                    for (int i = 0; i < invoiceList.Count; i++)
                    {
                        foreach (var category in data)
                        {
                            IInvoiceRet invoice = invoiceList.GetAt(i);
                            string? memo = invoice.Memo != null ? Convert.ToString(invoice.Memo.GetValue()) : null;
                            if (memo != null)
                            {
                                if (memo == $"{date}-{category.Key}")
                                {
                                    count++;
                                    string invoiceID = invoice.RefNumber.GetValue();
                                    string customerName = invoice.ClassRef.FullName.GetValue();
                                    string balanceAmount = invoice.BalanceRemaining.GetValue().ToString();
                                    decimal totalAmount = Convert.ToDecimal(invoice.Subtotal.GetValue());
                                    string txnDate = invoice.TxnDate.GetValue().ToString();

                                    Console.WriteLine($"Invoice ID: {invoiceID} | Customer: {customerName} | Balance Amount: {balanceAmount} | Paid Amount: {totalAmount} {txnDate}");

                                }
                            }
                        }

                    }
                    Console.WriteLine($"Invoices in QuickBooks {invoiceList.Count}:");

                    Console.WriteLine($"After Validation {count}:");
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
