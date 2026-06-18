const SPREADSHEET_ID = "1UPBi1Rm7Dw8-FFAC-eZs3MqzBk5RtmNuw5TMEQJ05Rg";

function getData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Get workers data
    const workersSheet = ss.getSheetByName("Workers") || ss.getSheets()[0];
    const workersData = workersSheet.getDataRange().getValues();
    
    // Get targets data
    const targetsSheet = ss.getSheetByName("Targets") || ss.getSheets()[1];
    const targetsData = targetsSheet.getDataRange().getValues();
    
    // Parse workers
    const workers = [];
    for(let i = 1; i < workersData.length; i++) {
      const row = workersData[i];
      workers.push({
        name: row[0] || "",
        review: row[1] || 0,
        video: row[2] || 0,
        mattPro: row[3] || 0,
        mattSel: row[4] || 0,
        recurring: row[5] || 0,
        daily: row[6] || 0,
        newSvc: row[7] || 0,
        pts: row[8] || 0,
        ranking: row[9] || "",
        photo: row[10] || "",
        actReward: row[11] || 0,
        bonusReward: row[12] || 0,
        total: row[13] || 0
      });
    }
    
    // Parse targets
    const targets = {};
    if(targetsData.length > 0) {
      const headers = targetsData[0];
      for(let i = 1; i < targetsData.length; i++) {
        const row = targetsData[i];
        targets[row[0]] = {
          target: row[1] || 0,
          semasa: row[2] || 0
        };
      }
    }
    
    return {
      workers: workers,
      targets: targets
    };
  } catch(e) {
    throw new Error("Error fetching data: " + e.message);
  }
}

function saveCustomer(customer) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Customers");
    
    // Create sheet if doesn't exist
    if(!sheet) {
      sheet = ss.insertSheet("Customers");
      const headers = ["ID", "Date", "Worker", "Category", "Name", "Phone", "Address", "Notes"];
      sheet.appendRow(headers);
    }
    
    // Append new customer
    const row = [
      customer.id,
      customer.date,
      customer.worker,
      customer.category,
      customer.name,
      customer.phone,
      customer.address,
      customer.notes
    ];
    sheet.appendRow(row);
    
    return { success: true, message: "Customer saved successfully" };
  } catch(e) {
    throw new Error("Error saving customer: " + e.message);
  }
}

function getCustomers() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Customers");
    
    if(!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const customers = [];
    
    for(let i = 1; i < data.length; i++) {
      const row = data[i];
      customers.push({
        id: row[0],
        date: row[1],
        worker: row[2],
        category: row[3],
        name: row[4],
        phone: row[5],
        address: row[6],
        notes: row[7]
      });
    }
    
    return customers.reverse(); // Latest first
  } catch(e) {
    throw new Error("Error fetching customers: " + e.message);
  }
}

function deleteCustomer(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Customers");
    
    if(!sheet) {
      throw new Error("Customers sheet not found");
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Find and delete row
    for(let i = 1; i < data.length; i++) {
      if(data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Customer deleted" };
      }
    }
    
    throw new Error("Customer not found");
  } catch(e) {
    throw new Error("Error deleting customer: " + e.message);
  }
}
