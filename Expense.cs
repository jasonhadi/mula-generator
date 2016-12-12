using System.Collections.Generic;
using MongoDB.Bson;

namespace mula_generator
{

    public class Receipt
    {
        public ObjectId parentExpense { get; set; }
        public ObjectId _id { get; set; }
        public ObjectId userId { get; set; }
        public string where { get; set; }
        public string type { get; set; }
        public double amount { get; set; }
        public string description { get; set; }
        public int __v { get; set; }
        public ObjectId parentProject { get; set; }
        public bool submitted { get; set; }
        public BsonDateTime lastUpdated { get; set; }
        public BsonDateTime created { get; set; }
        public BsonDateTime date { get; set; }
        public int sheetNumber { get; set; }
        public int projectNumber { get; set; }
        public int receiptNumber { get; set; }
    }

    public class Row
    {
        public ObjectId _id { get; set; }
        public int number { get; set; }
        public int sheetNumber { get; set; }
    }

    public class Project
    {
        public ObjectId _id { get; set; }
        public ObjectId userId { get; set; }
        public string assignment { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public int __v { get; set; }
        public ObjectId parentExpense { get; set; }
        public BsonDateTime lastUpdated { get; set; }
        public bool submitted { get; set; }
        public BsonDateTime created { get; set; }
        public List<Row> row { get; set; }
        public int receiptCount { get; set; }
    }

    public class Expense
    {
        public ObjectId _id { get; set; }
        public string fullname { get; set; }
        public string expCurrency { get; set; }
        public string reimbCurrency { get; set; }
        public int __v { get; set; }
        public ObjectId userId { get; set; }
        public BsonDateTime lastUpdated { get; set; }
        public BsonDateTime created { get; set; }
        public List<Receipt> receipts { get; set; }
        public List<Project> projects { get; set; }
        public bool submitted { get; set; }
        public int sheetCount { get; set; }
        public int receiptCount { get; set; }
        public BsonDateTime oldestBillDate { get; set; }
        public BsonDateTime submitDate { get; set; }
    }

}