using System;

namespace QuickBooksCRUD
{
    public class Program
    {

        public static void Main(string[] args)
        {
            QuickBooks quickBooks = new QuickBooks();
            //quickBooks.DoInvoiceAdd();
            //quickBooks.DoItemAdd();
            //quickBooks.GetAccount();
            //quickBooks.GetItems();
            //quickBooks.GetInvoices();
            quickBooks.GetCompanyInfo();

            quickBooks.DoInvoiceAdd();
        }

    }
}


