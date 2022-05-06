 function onUpdateClientData() {
  trace("onUpdateClientData");  
  HubSpot.getClientData();
}

//----------------------------------------------------------------------------------------
// Wrapper for https://developers.hubspot.com/docs/api/crm/understanding-the-crm
//   Contacts   : https://developers.hubspot.com/docs/api/crm/contacts
//   Deals      : https://developers.hubspot.com/docs/api/crm/deals

class HubSpot {

  static getUrl(method) {
    return `${HubSpot.baseUrl}/${method}&hapikey=${HubSpot.key}`; // for Contacts & Deals
  }

  static listContacts() {
    let url = HubSpot.getUrl("contacts?limit=100&properties=hs_object_id,firstname,lastname,createdate,email,hs_email_domain,phone,annualrevenue,how_many_people_in_total_including_the_couple_will_be_at_the_ceremony_and_or_the_celebration_,asana_link,hs_lifecyclestage_customer_date,hs_lifecyclestage_lead_date,hs_lifecyclestage_marketingqualifiedlead_date,hs_lifecyclestage_salesqualifiedlead_date,hs_lifecyclestage_subscriber_date,hs_lifecyclestage_evangelist_date,hs_lifecyclestage_opportunity_date,hs_lifecyclestage_other_date,city,company,hs_object_id,country,date,date_worked,do_you_agree_to_special_terms_in_the_event_of_a_coronavirus_event,hs_content_membership_email_confirmed,event_start_time,industry,is_there_any_food_that_you_dislike,is_your_kitchen_fulled_equipped_and_functional,jobtitle,kitchen,kitchen_1,lastmodifieddate,hs_latest_sequence_ended_date,hs_latest_sequence_enrolled,hs_latest_sequence_enrolled_date,lifecyclestage,hs_marketable_status,hs_marketable_reason_id,hs_marketable_reason_type,hs_marketable_until_renewal,mobilephone,numemployees,hs_sequences_enrolled_count,hs_createdate,hs_persona,zip,hs_language,salutation,state,address,hs_content_membership_registration_email_sent_at,time_sheet,twitterhandle,website,what_the_occasion&archived=false");
    let response = UrlFetchApp.fetch(url);
    trace(`HubSpot.listContacts --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
    let results = data['results'];
    let paging = data.paging.next;
    let sheet = SpreadsheetApp.getActiveSheet();
    let header = [ "ID", "First Name", "Last Name", "Create Date", "Email", "Email Domain", "Phone", "Annual Revenue", "Guests", "Asana Link", "Became a Customer Date", "Became a Lead Date", "Became a Marketing Qualified Lead Date", "Became a Sales Qualified Lead Date",  "Became a Subscriber Date",  "Became an Evangelist Date",  "Became an Opportunity Date",  "Became an Other Lifecycle Date", "City", "Company Name", "Contact ID", "Country/Region", "Date", "Date Worked", "Agree to Special Terms", "Email Confirmed", "Event Start Time", "Industry", "Is there any food you dislike", "Is your kitchen fulled equipped & functional", "Job Title", "Kitchen", "Kitchen 1", "Last Modified Date", "Last Sequence Ended Date", "Last Sequence Enrolled", "Last Sequence Enrolled Date", "Lifecycle Stage", "Marketing Contact Status", "Marketing Contact Status Source Name", "Marketing Contact Status Source Type", "Marketing Contact Until Next Update", "Mobile Phone Number", "Number of Employees", "Number of Sequences Enrolled", "Object Create Date/Time", "Persona", "Postal Code", "Preferred Language", "Salutation", "State/Region", "Street Address", "Time Registration Email Was Sent", "Time Sheet", "Twitter Username", "Website", "What the Occasion","Paging After","Paging Link"]; 
    let items = [header];
    results.forEach(function (result) {
    items.push([ result['properties'].hs_object_id, result['properties'].firstname, result['properties'].lastname, result['properties'].createdate, result['properties'].email, result['properties'].hs_email_domain, result['properties'].phone,  result['properties'].annualrevenue,  result['properties'].how_many_people_in_total_including_the_couple_will_be_at_the_ceremony_and_or_the_celebration_,  result['properties'].asana_link, result['properties'].hs_lifecyclestage_customer_date,  result['properties'].hs_lifecyclestage_lead_date, result['properties'].hs_lifecyclestage_marketingqualifiedlead_date,  result['properties'].hs_lifecyclestage_salesqualifiedlead_date,  result['properties'].hs_lifecyclestage_subscriber_date,  result['properties'].hs_lifecyclestage_evangelist_date,  result['properties'].hs_lifecyclestage_opportunity_date,  result['properties'].hs_lifecyclestage_other_date,  result['properties'].city,  result['properties'].company,  result['properties'].hs_object_id,  result['properties'].country,  result['properties'].date,  result['properties'].date_worked, result['properties'].do_you_agree_to_special_terms_in_the_event_of_a_coronavirus_event,  result['properties'].hs_content_membership_email_confirmed,  result['properties'].event_start_time,  result['properties'].industry,  result['properties'].is_there_any_food_that_you_dislike,  result['properties'].is_your_kitchen_fulled_equipped_and_functional,  result['properties'].jobtitle,  result['properties'].kitchen,  result['properties'].kitchen_1,  result['properties'].lastmodifieddate,  result['properties'].hs_latest_sequence_ended_date,  result['properties'].hs_latest_sequence_enrolled,  result['properties'].hs_latest_sequence_enrolled_date,  result['properties'].lifecyclestage,  result['properties'].hs_marketable_status,  result['properties'].hs_marketable_reason_id,  result['properties'].hs_marketable_reason_type,  result['properties'].hs_marketable_until_renewal,  result['properties'].mobilephone,  result['properties'].numemployees,  result['properties'].hs_sequences_enrolled_count, result['properties'].hs_createdate, result['properties'].hs_persona, result['properties'].zip, result['properties'].hs_language, result['properties'].salutation, result['properties'].state, result['properties'].address, result['properties'].hs_content_membership_registration_email_sent_at, result['properties'].time_sheet, result['properties'].twitterhandle, result['properties'].website, result['properties'].what_the_occasion, paging.after, paging.link]);
    });
    sheet.getRange(3,1,items.length,items[0].length).setValues(items);
  }

  static listDeals() {
    let url = HubSpot.getUrl("deals?limit=100&properties=hs_object_id,amount,closedate,createdate,dealname,description,hubspot_owner_id,dealstage,dealtype,departure_date,hs_forecast_amount,hs_manual_forecast_category,hs_forecast_probability,hubspot_team_id,hs_lastmodifieddate,hs_next_step,num_associated_contacts,hs_priority,pipeline&archived=false");
    let response = UrlFetchApp.fetch(url);
    //trace(`HubSpot.listDeals --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
    let results = data['results'];
    let paging = data.paging.next;
    let sheet = SpreadsheetApp.getActiveSheet();
    let header = ["Deal ID", "Amount", "Close Date", "Create Date", "Deal Name", "Deal Description", "Deal Owner", "Deal Type", "Deal Stage", "Departure Date", "Forecast Amount", "Forecast Category", "Forecast Probabilty", "HubSpot Team", "Last Modified Date", "Next Step", "Number of Contacts", "Priority", "Pipeline"];
    let items = [header];
    results.forEach(function (result) {
    if(result['properties'].dealstage !== "closedlost") {
      items.push([ result['properties'].hs_object_id, result['properties'].amount, result['properties'].closedate, result['properties'].createdate, result['properties'].dealname, result['properties'].description, result['properties'].hubspot_owner_id, result['properties'].dealtype, result['properties'].dealstage, result['properties'].departure_date, result['properties'].hs_forecast_amount, result['properties'].hs_manual_forecast_category, result['properties'].hs_forecast_probability, result['properties'].hubspot_team_id, result['properties'].hs_lastmodifieddate, result['properties'].hs_next_step, result['properties'].num_associated_contacts, result['properties'].priority, result['properties'].pipeline]);      
      }
    });
    /*let apiCall = function(url) {
      let response = UrlFetchApp.fetch(url);
      let data = JSON.parse(response);
      return data;
    };
    apiCall(paging.link);*/
    trace(`Paging After : ${ paging.after}, Link : ${ paging.link+"&hapikey=0020bf99-6b2a-4887-90af-adac067aacba" }`);
    sheet.getRange(3,1,items.length,items[0].length).setValues(items);
    //if (paging) { return true};
  }

  static listEngagement() {
    let url = "https://api.hubapi.com/engagements/v1/engagements/paged?hapikey=0020bf99-6b2a-4887-90af-adac067aacba&limit=250";
    let response = UrlFetchApp.fetch(url);
    trace(`HubSpot.listEngagement --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
    let results = data['results'];
    let sheet = SpreadsheetApp.getActiveSheet();
    let header = ["ID", "Portal ID", "Active", "Created At", "Last Updated", "Created By", "Modified By", "Owner ID", "Type", "Timestamp", "Source", "Source ID", "Accessible Team IDs", "Queue Membership IDs", "Body Preview Is Truncated", "gdpr Deleted", "Contact IDs", "Deal IDs", "Status", "Object Type", "Title", "Task Type", "Reminders", "Send Default Reminder", "Priority", "Is All Day", "Completion Date"];
    let items = [header];
    results.forEach(function (result) {
      items.push([ result['engagement'].id, result['engagement'].portalId, result['engagement'].active, result['engagement'].createdAt, result['engagement'].lastUpdated,result['engagement'].createdBy, result['engagement'].modifiedBy, result['engagement'].ownerId, result['engagement'].type, result['engagement'].timestamp, result['engagement'].source, result['engagement'].sourceId, result['engagement'].allAccessibleTeamIds, result['engagement'].queueMembershipIds, result['engagement'].bodyPreviewIsTruncated, result['engagement'].gdprDeleted, result['associations'].contactIds, result['associations'].dealIds, result['metadata'].status, result['metadata'].forObjectType, result['metadata'].subject, result['metadata'].taskType, result['metadata'].reminders, result['metadata'].sendDefaultReminder, result['metadata'].priority, result['metadata'].isAllDay, result['metadata'].completionDate]);
    });    
    sheet.getRange(3,1,items.length,items[0].length).setValues(items);
  }

  static masterHubspot() {
    let url = HubSpot.getUrl("contacts?limit=100&properties=invoice_no,first_conversion_date,deal_status,hubspot_owner_id,contacttype,number_of_guests__inc_the_couple_,decor,firstname,lastname,partner_s_first_name,partner_s_last_name,meet___greet_date,meet___greet_time,event_date,time_of_ceremony,confirmed_venue,venue___reception,do_you_require_witnesses_,documents_status,notes,registry_office_payment,xero_outstanding_thirty_days,florist,suppliers_deposits&archived=false");
    let response = UrlFetchApp.fetch(url);
    trace(`HubSpot.masterHubspot --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
    let results = data['results'];
    let paging = data.paging.next;
    let sheet = SpreadsheetApp.getActiveSheet();
    let header = ["Invoice Number", "Enquiry Date", "Status", "WP", "Type of Event", "# of Guests", "Decor Required", "Client Name", "Partner Full Name", "Meet & Greet Date", "Meet & Greet Time", "Event Date", "Event Time", "Venue 1", "Venue 2", "Witnesses", "Registry Office Documents", "Notes","Registry Amount Paid", "Outstanding", "Florist", "Suppliers Costings"];
    let items = [header];
    results.forEach(function (result) {
    items.push([ result['properties'].invoice_no, result['properties'].first_conversion_date, result['properties'].deal_status, result['properties'].hubspot_owner_id, result['properties'].contacttype, result['properties'].number_of_guests__inc_the_couple_, result['properties'].decor, result['properties'].firstname +" "+result['properties'].lastname, result['properties'].partner_s_first_name +" "+ result['properties'].partner_s_last_name, result['properties'].meet___greet_date, result['properties'].meet___greet_time, result['properties'].event_date, result['properties'].time_of_ceremony, result['properties'].confirmed_venue, result['properties'].venue___reception, result['properties'].do_you_require_witnesses_, result['properties'].documents_status, result['properties'].notes, result['properties'].registry_office_payment, result['properties'].xero_outstanding_thirty_days, result['properties'].florist, result['properties'].suppliers_deposits]);
    });
    sheet.getRange(3,1,items.length,items[0].length).setValues(items);
  }  

  static contactToDeal() {
    let url = "https://api.hubapi.com/crm/v3/objects/contacts?associations=deal&hapikey=0020bf99-6b2a-4887-90af-adac067aacba";
    let response = UrlFetchApp.fetch(url);
    trace(`HubSpot.contactToDeal --> ${response.getContentText()}`);
    let data = JSON.parse(response.getContentText());
    let results = data['results'];
    results.forEach(function (result) {
      Logger.log(result['associations']);
    });
  }

  static dealToContact(dealId) {
    trace("dealToContact")
    let url = "https://api.hubapi.com/crm/v3/objects/deals/"+dealId+"/associations/contacts?hapikey=0020bf99-6b2a-4887-90af-adac067aacba"
    let response = UrlFetchApp.fetch(url);
    let data = JSON.parse(response.getContentText());
    let result = Array.from(data['results']);
    let contactId, i = 0;
    let rl = result.length;
    for(i; i < rl; ++i) {
      if( result[i].type == "deal_to_contact")
      {
        contactId = result[i].id;        
        trace(`Associated Contact ID : ${contactId}`);
        HubSpot.listContact(contactId);
      }
    }
  }

  static getClientData() {
    trace("getClientData");
    //try {
      let dealId = Spreadsheet.getCellValue("HubSpotDeal");
      trace(`Deal ID : ${dealId}`);
      Dialog.notify("Updating Client Data...","Please wait updating Client Data from HubSpot!");
      let url = "https://api.hubapi.com/crm/v3/objects/deals/"+dealId+"?properties=hs_object_id,invoice_id,estimate_id,amount,closedate,createdate,dealname,description,hubspot_owner_id,dealstage,dealtype,departure_date,hs_forecast_amount,hs_manual_forecast_category,hs_forecast_probability,hubspot_team_id,hs_lastmodifieddate,hs_next_step,num_associated_contacts,hs_priority,pipeline&archived=false&hapikey=0020bf99-6b2a-4887-90af-adac067aacba";
      let response = UrlFetchApp.fetch(url);
      let data = JSON.parse(response.getContentText());
      console.log(data);
      let result = data;
      let range = SpreadsheetApp.getActive().getRangeByName("DealData");

      let props = ["invoice_id","estimate_id","amount","closedate","createdate","dealname","description","hubspot_owner_id","dealtype","dealstage","departure_date","hs_forecast_amount","hs_manual_forecast_category","hs_forecast_probability","hubspot_team_id","hs_lastmodifieddate","hs_next_step","num_associated_contacts","hs_priority","pipeline" ]; // 20 deal properties

      let object = [];                                // For source object element

      let header = ["Invoice ID", "Estimate ID", "Amount", "Close Date", "Create Date", "Deal Name", "Deal Description", "Deal Owner", "Deal Type", "Deal Stage", "Departure Date", "Forecast Amount", "Forecast Category", "Forecast Probabilty", "HubSpot Team", "Last Modified Date", "Next Step", "Number of Contacts", "Priority", "Pipeline"];        // header for the properties    

      let items = [props,object,header];   
      items.push([ result['properties'].invoice_id, result['properties'].estimate_id, result['properties'].amount, result['properties'].closedate, result['properties'].createdate, result['properties'].dealname, result['properties'].description, result['properties'].hubspot_owner_id, result['properties'].dealtype, result['properties'].dealstage, result['properties'].departure_date, result['properties'].hs_forecast_amount, result['properties'].hs_manual_forecast_category, result['properties'].hs_forecast_probability, result['properties'].hubspot_team_id, result['properties'].hs_lastmodifieddate, result['properties'].hs_next_step, result['properties'].num_associated_contacts, result['properties'].priority, result['properties'].pipeline]);
      range.setValues(Object.keys(items[0]).map( function(columnNumber) {
        return items.map( function(row) { 
          object.push(["Deal"]);
          return row[columnNumber];
        });
      }));
      HubSpot.dealToContact(dealId);
    //}
    /*catch {
      Dialog.notify("HubSpot Deal Not Found!","Please enter the respective client's Deal ID in HubSpot Deal!");
    }*/
  }


  static listContact(contactId) {
    trace("listContact");
    //try {
      let url = "https://api.hubapi.com/crm/v3/objects/contacts/"+contactId+"?properties=invoice_no,first_conversion_date,deal_status,hubspot_owner_id,contacttype,number_of_guests__inc_the_couple_,decor,hs_object_id,firstname,lastname,meet___greet_date,meet___greet_time,event_date,confirmed_wedding_date,time_of_ceremony,confirmed_venue,venue___reception,do_you_require_witnesses_,documents_status,notes,createdate,email,hs_email_domain,phone,annualrevenue,number_of_guests,asana_link,hs_lifecyclestage_customer_date,hs_lifecyclestage_lead_date,hs_lifecyclestage_marketingqualifiedlead_date,hs_lifecyclestage_salesqualifiedlead_date,hs_lifecyclestage_subscriber_date,hs_lifecyclestage_evangelist_date,hs_lifecyclestage_opportunity_date,hs_lifecyclestage_other_date,city,company,hs_object_id,country,date,date_worked,do_you_agree_to_special_terms_in_the_event_of_a_coronavirus_event,hs_content_membership_email_confirmed,event_start_time,industry,is_there_any_food_that_you_dislike,is_your_kitchen_fulled_equipped_and_functional,jobtitle,kitchen,kitchen_1,lastmodifieddate,hs_latest_sequence_ended_date,hs_latest_sequence_enrolled,hs_latest_sequence_enrolled_date,lifecyclestage,hs_marketable_status,hs_marketable_reason_id,hs_marketable_reason_type,hs_marketable_until_renewal,mobilephone,numemployees,hs_sequences_enrolled_count,hs_createdate,hs_persona,zip,hs_language,salutation,state,address,hs_content_membership_registration_email_sent_at,time_sheet,twitterhandle,website,what_the_occasion&archived=false&hapikey=0020bf99-6b2a-4887-90af-adac067aacba";
      let response = UrlFetchApp.fetch(url);
      let data = JSON.parse(response.getContentText());
      console.log(data);
      let result =data;
      let range = SpreadsheetApp.getActive().getRangeByName("ContactData");

      let props = ["invoice_no","first_conversion_date","deal_status","hubspot_owner_id","contacttype","number_of_guests__inc_the_couple_","decor","firstname","lastname","meet___greet_date","meet___greet_time","event_date","confirmed_wedding_date","time_of_ceremony","confirmed_venue","venue___reception","do_you_require_witnesses_","documents_status","notes","createdate","email","hs_email_domain","phone","annualrevenue","number_of_guests","asana_link","hs_lifecyclestage_customer_date","hs_lifecyclestage_lead_date","hs_lifecyclestage_marketingqualifiedlead_date","hs_lifecyclestage_salesqualifiedlead_date","hs_lifecyclestage_subscriber_date","hs_lifecyclestage_evangelist_date","hs_lifecyclestage_opportunity_date","hs_lifecyclestage_other_date","city","company","hs_object_id","country","date","date_worked","do_you_agree_to_special_terms_in_the_event_of_a_coronavirus_event","hs_content_membership_email_confirmed","event_start_time","industry","is_there_any_food_that_you_dislike","is_your_kitchen_fulled_equipped_and_functional","jobtitle","kitchen","kitchen_1","lastmodifieddate","hs_latest_sequence_ended_date","hs_latest_sequence_enrolled","hs_latest_sequence_enrolled_date","lifecyclestage","hs_marketable_status","hs_marketable_reason_id","hs_marketable_reason_type","hs_marketable_until_renewal","mobilephone","numemployees","hs_sequences_enrolled_count","hs_createdate","hs_persona","zip","hs_language","salutation","state","address","hs_content_membership_registration_email_sent_at","time_sheet","twitterhandle","website","what_the_occasion"];     // 73 contact properties   

      let object = [];                              // For source object element    

      let header = ["Invoice No.", "Enquiry Date", "Status", "WP", "Type of Event", "# of Guests", "Decor Required", "First Name", "Last Name","Meet & Greet Date", "Meet & Greet Time", "Event Date","Confirmed wedding date", "Event Time", "Venue 1", "Venue 2", "Witnesses", "Registry Office Documents", "Notes", "Create Date", "Email", "Email Domain", "Phone", "Annual Revenue", "Guests", "Asana Link", "Became a Customer Date", "Became a Lead Date", "Became a Marketing Qualified Lead Date", "Became a Sales Qualified Lead Date",  "Became a Subscriber Date",  "Became an Evangelist Date",  "Became an Opportunity Date",  "Became an Other Lifecycle Date", "City", "Company Name", "Contact ID", "Country/Region", "Date", "Date Worked", "Agree to Special Terms", "Email Confirmed", "Event Start Time", "Industry", "Is there any food you dislike", "Is your kitchen fulled equipped & functional", "Job Title", "Kitchen", "Kitchen 1", "Last Modified Date", "Last Sequence Ended Date", "Last Sequence Enrolled", "Last Sequence Enrolled Date", "Lifecycle Stage", "Marketing Contact Status", "Marketing Contact Status Source Name", "Marketing Contact Status Source Type", "Marketing Contact Until Next Update", "Mobile Phone Number", "Number of Employees", "Number of Sequences Enrolled", "Object Create Date/Time", "Persona", "Postal Code", "Preferred Language", "Salutation", "State/Region", "Street Address", "Time Registration Email Was Sent", "Time Sheet", "Twitter Username", "Website", "What the Occasion"];    // header for the properties 
      
      let items = [props,object,header];
      items.push([ result['properties'].invoice_no, result['properties'].first_conversion_date, result['properties'].deal_status, result['properties'].hubspot_owner_id, result['properties'].contacttype, result['properties'].number_of_guests__inc_the_couple_, result['properties'].decor, result['properties'].firstname, result['properties'].lastname,result['properties'].meet___greet_date, result['properties'].meet___greet_time, result['properties'].event_date, result['properties'].confirmed_wedding_date, result['properties'].time_of_ceremony, result['properties'].confirmed_venue, result['properties'].venue___reception, result['properties'].do_you_require_witnesses_, result['properties'].documents_status, result['properties'].notes, result['properties'].createdate, result['properties'].email, result['properties'].hs_email_domain, result['properties'].phone,  result['properties'].annualrevenue, result['properties'].number_of_guests,  result['properties'].asana_link, result['properties'].hs_lifecyclestage_customer_date,  result['properties'].hs_lifecyclestage_lead_date, result['properties'].hs_lifecyclestage_marketingqualifiedlead_date,  result['properties'].hs_lifecyclestage_salesqualifiedlead_date,  result['properties'].hs_lifecyclestage_subscriber_date, result['properties'].hs_lifecyclestage_evangelist_date,  result['properties'].hs_lifecyclestage_opportunity_date, result['properties'].hs_lifecyclestage_other_date, result['properties'].city,  result['properties'].company,  result['properties'].hs_object_id,  result['properties'].country, result['properties'].date,  result['properties'].date_worked, result['properties'].do_you_agree_to_special_terms_in_the_event_of_a_coronavirus_event, result['properties'].hs_content_membership_email_confirmed,  result['properties'].event_start_time,  result['properties'].industry, result['properties'].is_there_any_food_that_you_dislike,  result['properties'].is_your_kitchen_fulled_equipped_and_functional, result['properties'].jobtitle,  result['properties'].kitchen,  result['properties'].kitchen_1,  result['properties'].lastmodifieddate,  result['properties'].hs_latest_sequence_ended_date, result['properties'].hs_latest_sequence_enrolled,  result['properties'].hs_latest_sequence_enrolled_date, result['properties'].lifecyclestage, result['properties'].hs_marketable_status,  result['properties'].hs_marketable_reason_id, result['properties'].hs_marketable_reason_type, result['properties'].hs_marketable_until_renewal,  result['properties'].mobilephone, result['properties'].numemployees, result['properties'].hs_sequences_enrolled_count, result['properties'].hs_createdate, result['properties'].hs_persona, result['properties'].zip, result['properties'].hs_language, result['properties'].salutation, result['properties'].state, result['properties'].address, result['properties'].hs_content_membership_registration_email_sent_at, result['properties'].time_sheet, result['properties'].twitterhandle, result['properties'].website, result['properties'].what_the_occasion]);
      range.setValues(Object.keys(items[0]).map ( function (columnNumber) {
        return items.map( function (row) {
          object.push(["Contact"]);
          return row[columnNumber];
        });
      }));
    //}
    /*catch {
      Dialog.notify("foo","bar");
    }*/
  }
}

class HubSpotDataDictionary {

  constructor() {
    trace("> NEW HubSpotDataDictionary, loading dictionary...");
    let dataDictionarySheet = Spreadsheet.openById(HubSpot.dataDictionarySheetId);
    this.summaryContactPropertyIds = dataDictionarySheet.getRangeByName("SummaryContactPropertyIds").values;
    this.summaryDealPropertyIds = dataDictionarySheet.getRangeByName("SummaryDealPropertyIds").values;
    this.detailedContactPropertyIds = dataDictionarySheet.getRangeByName("DetailedContactPropertyIds").values;
    this.detailedDealPropertyIds = dataDictionarySheet.getRangeByName("DetailedDealPropertyIds").values;

    this.detailedContactPropertyIds.forEach( item => { if (item[0] !== "") trace(`item: ${item[0]}`); } );
    trace("< NEW HubSpotDataDictionary, dictionary loaded.");
  }

  static get current() {
    return HubSpotDataDictionary.singleton ?? (HubSpotDataDictionary.singleton = new HubSpotDataDictionary);
  }

}

HubSpot.baseUrl = "https://api.hubapi.com/crm/v3/objects";
HubSpot.key = "0020bf99-6b2a-4887-90af-adac067aacba";
HubSpot.dataDictionarySheetId = "1C_uOMH30siZLSGzYOziBfcl_lx5ZZRYCBv7fqpegGEk";

