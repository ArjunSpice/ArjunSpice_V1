ARJUN SPICE - DELIVERY AND EMAIL UPDATE

1. BACK UP your spreadsheet and current Apps Script project first.
2. Replace Code.gs with the included Code.gs. Replace Index.html only if needed; this ZIP includes the same HTML/images from your prior working package.
3. Project Settings > Script Properties: SPREADSHEET_ID and OWNER_EMAIL=arjunspice8@gmail.com.
4. Save and run setupArjunSpice from Apps Script editor. Authorize spreadsheet and MailApp permissions.
5. In Google Sheet > Orders > Admin Action dropdown, choose Confirm Order. Confirmation email is sent before status changes to Confirmed. Check Customer Confirmation Email Status.
6. When delivered, choose Mark Delivered in Admin Action (or set Status to Delivered). A delivery email is attempted. On success, order is moved to Delivered Orders and removed from Orders.
7. To START FRESH and remove ALL existing order details from active sheets, separately run archiveAndClearOldOrders ONCE from Apps Script editor. This backs up all prior Orders, Order Items, and Delivered Orders as hidden timestamped sheets before clearing them. Catalog and Settings are preserved. Do not run this after accepting new orders you want to keep.
8. Deploy > Manage deployments > Edit > New version > Deploy. Changing product prices later does not require redeployment.
9. Test: submit a COD order with a valid customer email; confirm; inspect confirmation email; mark delivered; inspect delivered email and archive row.

IMPORTANT: Existing orders may have no customer email, or an incorrectly entered email. Those cannot receive confirmation. Check the Email cell and Customer Confirmation Email Status. Apps Script > Executions shows trigger errors. Google MailApp quotas and authorization apply. A 'Sent' result indicates MailApp accepted the message, not guaranteed inbox delivery; check spam too.

For a standalone Apps Script project, a custom Google Sheets menu may not appear; the Admin Action dropdown is the supported control.
