using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using MainStreet.BusinessFlow.SDK.Ws;
using MainStreet.BusinessFlow.SDK.Web;
using MainStreet.BusinessFlow.SDK;
using System.IO;

namespace ChannelSale
{
    class Program
    {
        static void Main(string[] args)
        {
            ChannelSaleFeed objChannel = null;
            try
            {
                if (IsValidConnection())
                {
                    Console.WriteLine("==========================================================");
                    Console.WriteLine("     CHANNEL SALE FEED ");
                    Console.WriteLine("==========================================================");
                    objChannel = new ChannelSaleFeed();
                    objChannel.ChannelDownloadFull();
                }
            }
            catch
            {
            }
        }

        private static bool IsValidConnection()
        {
            //return true;
            try
            {
                OrderSearchRequest osr = new OrderSearchRequest();
                osr.MaxRows = 1;
                dsOrderList dsOrders = BusinessFlow.WebServices.Order.Search(osr);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}