try {

    var dealId = $Page.record_id;

    if (!dealId) {
        ZDK.Client.showMessage('Deal ID not found. Open this from a Deal detail page.', {
            type: 'error'
        });
        return;
    }

    var headers = {
        'Content-Type': 'application/json'
    };

    // fields param is required to bring back the subform (Product_Details) and
    // related fields — they are not returned by default on a plain GET.
    var dealFields = [
        'Deal_Name',
        'Product_Details',
        'Payment_Profile_Table',
        'Contact_Name',
        'Account_Name',
        'Lead_Name',
        'Owner'
    ].join(',');

    var dealResponse = ZDK.Apps.CRM.Connections.invoke(
        'crm',
        'https://www.zohoapis.com/crm/v6/Deals/' + dealId + '?fields=' + dealFields,
        'GET',
        2,
        {},
        headers
    );

    console.log('Deal fetch response:', dealResponse);

    var deal = null;

    try {
        var dealBody = dealResponse._details.statusMessage;
        if (typeof dealBody === 'string') {
            dealBody = JSON.parse(dealBody);
        }
        deal = dealBody.data[0];
    } catch (e) {
        console.log('Deal parse error:', e);
    }

    if (!deal) {
        ZDK.Client.showMessage('Could not load Deal record.', {
            type: 'error'
        });
        return;
    }

    var productDetails = deal.Product_Details || [];

    if (productDetails.length === 0) {
        ZDK.Client.showMessage('Please select at least one product to convert a Deal into a Quote.', {
            type: 'error'
        });
        return;
    }

    // ----- Build Quoted_Items from Product_Details -----
    var itemList = [];

    for (var i = 0; i < productDetails.length; i++) {
        var item = productDetails[i];

        itemList.push({
            Product_Name: (item.Product_Name && item.Product_Name.id) ? item.Product_Name.id : '',
            Description: item.Description || '',
            List_Price: (item.List_Price != null) ? item.List_Price : 0,
            Quantity: item.Quantity,
            Discount: (item.Discount2 != null) ? item.Discount2 : 0,
            Start_Date: item.Start_Date || '',
            End_Date: item.End_Date || '',
            Product_Line_Type: item.Product_Line_Type || '',
            Comments: item.Comments || ''
        });
    }

    // ----- Build Payment_Profile_Table1 from Payment_Profile_Table -----
    var paymentList = [];
    var paymentProfileTable = deal.Payment_Profile_Table || [];

    for (var j = 0; j < paymentProfileTable.length; j++) {
        var row = paymentProfileTable[j];

        paymentList.push({
            Due_Date: row.Due_Date,
            Amount_Due: row.Amount_Due
        });
    }

    var quoteRecord = {
        Subject: deal.Deal_Name,
        Deal_Name: { id: String(dealId) },
        Quoted_Items: itemList,
        Payment_Profile_Table1: paymentList,
        Contact_Name: deal.Contact_Name,
        Account_Name: deal.Account_Name,
        Owner: $Crm.user.id,
        Quote_Stage: 'Draft'
    };

    // ----- Shipping / Billing logic based on the linked Account -----
    if (deal.Account_Name && deal.Account_Name.id) {

        var accountFields = [
            'Shipping_Account',
            'Shipping_Contact',
            'Shipping_Street',
            'Shipping_City',
            'Shipping_State',
            'Shipping_Code',
            'Shipping_Country',
            'Billing_Street',
            'Billing_City',
            'Billing_State',
            'Billing_Code',
            'Billing_Country'
        ].join(',');

        var accountResponse = ZDK.Apps.CRM.Connections.invoke(
            'crm',
            'https://www.zohoapis.com/crm/v6/Accounts/' + deal.Account_Name.id + '?fields=' + accountFields,
            'GET',
            2,
            {},
            headers
        );

        console.log('Account fetch response:', accountResponse);

        var account = null;

        try {
            var accountBody = accountResponse._details.statusMessage;
            if (typeof accountBody === 'string') {
                accountBody = JSON.parse(accountBody);
            }
            account = accountBody.data[0];
        } catch (e) {
            console.log('Account parse error:', e);
        }

        if (account) {
            if (account.Shipping_Account) {
                quoteRecord.Shipping_Contact = account.Shipping_Contact;
                quoteRecord.Reseller_Account = account.Shipping_Account;
                quoteRecord.Billing_Street = account.Shipping_Street;
                quoteRecord.Billing_City = account.Shipping_City;
                quoteRecord.Billing_State = account.Shipping_State;
                quoteRecord.Billing_Code = account.Shipping_Code;
                quoteRecord.Billing_Country = account.Shipping_Country;
                quoteRecord.Shipping_Street = account.Billing_Street;
                quoteRecord.Shipping_City = account.Billing_City;
                quoteRecord.Shipping_State = account.Billing_State;
                quoteRecord.Shipping_Code = account.Billing_Code;
                quoteRecord.Shipping_Country = account.Billing_Country;
            } else {
                quoteRecord.Shipping_Contact = deal.Contact_Name;
                quoteRecord.Reseller_Account = deal.Account_Name;
                quoteRecord.Billing_Street = account.Billing_Street;
                quoteRecord.Billing_City = account.Billing_City;
                quoteRecord.Billing_State = account.Billing_State;
                quoteRecord.Billing_Code = account.Billing_Code;
                quoteRecord.Billing_Country = account.Billing_Country;
                quoteRecord.Shipping_Street = account.Shipping_Street;
                quoteRecord.Shipping_City = account.Shipping_City;
                quoteRecord.Shipping_State = account.Shipping_State;
                quoteRecord.Shipping_Code = account.Shipping_Code;
                quoteRecord.Shipping_Country = account.Shipping_Country;
            }
        }
    }

    var createQuotePayload = {
        data: [quoteRecord]
    };

    console.log('Create Quote API payload:', JSON.stringify(createQuotePayload));

    var createQuoteResponse = ZDK.Apps.CRM.Connections.invoke(
        'crm',
        'https://www.zohoapis.com/crm/v6/Quotes',
        'POST',
        2,
        createQuotePayload,
        headers
    );

    console.log('Create Quote API response:', createQuoteResponse);

    var quoteId = '';

    try {
        var body = createQuoteResponse._details.statusMessage;
        if (typeof body === 'string') {
            body = JSON.parse(body);
        }
        quoteId = body.data[0].details.id;
    } catch (e) {
        console.log(e);
    }

    if (!quoteId) {
        ZDK.Client.showMessage('Quote creation failed. Check console logs.', {
            type: 'error'
        });
        return;
    }

    // ----- Verify Quote_Stage actually persisted on the created Quote -----
    try {
        var verifyQuoteResponse = ZDK.Apps.CRM.Connections.invoke(
            'crm',
            'https://www.zohoapis.com/crm/v6/Quotes/' + quoteId + '?fields=Quote_Stage',
            'GET',
            2,
            {},
            headers
        );

        var verifyBody = verifyQuoteResponse._details.statusMessage;
        if (typeof verifyBody === 'string') {
            verifyBody = JSON.parse(verifyBody);
        }

        console.log('Verify Quote_Stage after creation:', JSON.stringify(verifyBody.data[0]));
    } catch (e) {
        console.log('Verify Quote_Stage error:', e);
    }

    // ----- NEW: update Deal Stage to Quote -----
    var updateDealPayload = {
        data: [
            {
                Stage: 'Quote'
            }
        ]
    };

    var updateDealResponse = ZDK.Apps.CRM.Connections.invoke(
        'crm',
        'https://www.zohoapis.com/crm/v6/Deals/' + dealId,
        'PUT',
        2,
        updateDealPayload,
        headers
    );

    console.log('Update Deal Stage API response:', updateDealResponse);
    // ----- END NEW -----

    // ----- Update Lead Status to Quote, if this Deal came from a Lead -----
    // Get the Lead Id from the Lead_Name lookup field on the Deal.
    var leadIdFromDeal = '';

    try {
        var leadNameField = deal.Lead_Name;

        if (leadNameField) {
            if (typeof leadNameField === 'string' || typeof leadNameField === 'number') {
                leadIdFromDeal = String(leadNameField);
            } else if (leadNameField.id) {
                leadIdFromDeal = String(leadNameField.id);
            }
        }
    } catch (e) {
        console.log('Lead_Name parse error:', e);
    }

    console.log('Resolved leadIdFromDeal:', leadIdFromDeal);

    if (leadIdFromDeal) {

        var updateLeadPayload = {
            data: [
                {
                    Lead_Status: 'Sales Admin'
                }
            ]
        };

        var updateLeadResponse = ZDK.Apps.CRM.Connections.invoke(
            'crm',
            'https://www.zohoapis.com/crm/v6/Leads/' + leadIdFromDeal,
            'PUT',
            2,
            updateLeadPayload,
            headers
        );

        console.log('Update Lead API response:', updateLeadResponse);
    } else {
        console.log('No Lead_Name on this Deal, skipping Lead_Status update.');
    }

    ZDK.Client.showMessage('Quote created successfully!!', {
        type: 'success'
    });
    ZDK.Client.navigateTo(
        "record_detail",
        {
            module: "Quotes",
            record_id: quoteId
        },
        "_self"
    );

} catch (e) {
    console.log('Convert to Quote error:', e);

    ZDK.Client.showMessage('Failed. Check console logs.', {
        type: 'error'
    });
}