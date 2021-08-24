using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;


	public class ContractItemsInputNew
	{

		public string uniquekey { get; set; }
		public string BatchNum { get; set; }
		public string BatchConNum { get; set; }
		public string Description { get; set; }
		public string Cost { get; set; }
		public string MFG { get; set; }
		public string Model { get; set; }
		public string SerialNum { get; set; }
		public string Quantity { get; set; }
		public string SeqNum { get; set; }
		public string LastUpdate { get; set; }
		public string sItemID { get; set; }
		public string sUniTran { get; set; }
		public string lastUser { get; set; }

		public ContractItemsInputNew(string vOption = "")
		{
				InitFields();
		}

		public void InitFields()
		{
			uniquekey = "";
			BatchNum = "";
			BatchConNum = "";
			Description = "";
			Cost = "";
			MFG = "";
			Model = "";
			SerialNum = "";
			Quantity = "";
			SeqNum = "";
			LastUpdate = "";
			sItemID = "";
			sUniTran = "";
			lastUser = "";
		}
	}