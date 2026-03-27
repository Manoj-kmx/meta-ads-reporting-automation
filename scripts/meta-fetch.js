function fetchMetaAdsInsights() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meta Insights");

  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Meta Insights");
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Date", "Campaign Name", "Ad Set Name", "Ad Name",
      "Impressions", "Reach", "Cost per Reach",
      "Clicks", "CPC", "CTR",
      "Add to Cart", "Cost per Add to Cart",
      "Purchase", "Cost per Purchase", "Purchase Value", "Purchase ROAS",
      "Checkout Initiated", "Cost per Checkout",
      "View Content", "Landing Page Views",
      "Add to Wishlist", "Search", "Subscribe", "Lead", "Start Trial",
      "Spend"
    ]);
  }

  // 🔐 Secure token (DO NOT hardcode)
  var accessToken = PropertiesService.getScriptProperties().getProperty("META_ACCESS_TOKEN");

  var baseURL = "https://graph.facebook.com/v22.0/act_YOUR_AD_ACCOUNT_ID/insights";
  var fields = "date_stop,campaign_name,adset_name,ad_name,impressions,reach,clicks,ctr,cpc,actions,action_values,spend,website_purchase_roas";

  var params = "&level=ad&time_range={'since':'2024-01-01','until':'2024-12-31'}&time_increment=1";

  var url = baseURL + "?access_token=" + accessToken + "&fields=" + fields + params;

  var data = [];

  while (url) {
    try {
      Logger.log("Fetching: " + url);

      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var json = JSON.parse(response.getContentText());

      if (json.error) {
        Logger.log("Error: " + json.error.message);
        return;
      }

      json.data.forEach(function(row) {
        var spend = parseFloat(row.spend || 0);
        var reach = parseFloat(row.reach || 0);
        var impressions = parseFloat(row.impressions || 0);
        var clicks = parseFloat(row.clicks || 0);
        var ctr = parseFloat(row.ctr || 0);
        var cpc = parseFloat(row.cpc || 0);
        var costPerReach = reach ? spend / reach : 0;

        var actionsMap = {};
        var valuesMap = {};

        (row.actions || []).forEach(function(action) {
          actionsMap[action.action_type] = parseFloat(action.value || 0);
        });

        (row.action_values || []).forEach(function(value) {
          valuesMap[value.action_type] = parseFloat(value.value || 0);
        });

        var addToCart = actionsMap["add_to_cart"] || 0;
        var costPerAddToCart = addToCart ? spend / addToCart : 0;

        var purchase = actionsMap["purchase"] || 0;
        var costPerPurchase = purchase ? spend / purchase : 0;
        var purchaseValue = valuesMap["purchase"] || 0;
        var purchaseROAS = row.website_purchase_roas?.[0]?.value || 0;

        var checkoutInitiated = actionsMap["checkout_initiated"] || 0;
        var costPerCheckout = checkoutInitiated ? spend / checkoutInitiated : 0;

        var viewContent = actionsMap["view_content"] || 0;
        var landingPageViews = actionsMap["landing_page_view"] || 0;

        var addToWishlist = actionsMap["add_to_wishlist"] || 0;
        var search = actionsMap["search"] || 0;
        var subscribe = actionsMap["subscribe"] || 0;
        var lead = actionsMap["lead"] || 0;
        var startTrial = actionsMap["start_trial"] || 0;

        data.push([
          row.date_stop, row.campaign_name, row.adset_name, row.ad_name,
          impressions, reach, costPerReach,
          clicks, cpc, ctr,
          addToCart, costPerAddToCart,
          purchase, costPerPurchase, purchaseValue, purchaseROAS,
          checkoutInitiated, costPerCheckout,
          viewContent, landingPageViews,
          addToWishlist, search, subscribe, lead, startTrial,
          spend
        ]);
      });

      url = json.paging?.next || null;
      Utilities.sleep(1500);

    } catch (error) {
      Logger.log("Fetch error: " + error.message);
      break;
    }
  }

  if (data.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
  }
}
