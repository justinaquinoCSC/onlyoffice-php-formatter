# PHP Currency Plugin for OnlyOffice

A plugin that converts and formats currencies to Philippine Peso (₱) with live exchange rates.

**Built by Claude Code** (Anthropic's Claude AI)

---

## Features

- **Live Exchange Rates**: Fetches current rates from frankfurter.app API
- **Currency Conversion**: Convert values from 20+ currencies to PHP
- **Currency Formatting**: Format cells as Philippine Peso (₱#,##0.00)
- **Accounting Format**: Negative values displayed in parentheses

## Supported Currencies

USD, EUR, GBP, JPY, AUD, CAD, CHF, CNY, HKD, SGD, KRW, THB, MYR, IDR, INR, NZD, MXN, BRL, AED, SAR, and more.

---

## Code Breakdown

### File Structure

```
onlyofficePH-260131/
├── config.json          # Plugin configuration
├── index.html           # UI with dropdown and buttons
├── scripts/
│   └── code.js          # Main plugin logic
├── resources/
│   ├── icon.png         # 40x40 plugin icon
│   └── icon@2x.png      # 80x80 retina icon
└── README.md            # This file
```

### config.json

Defines plugin metadata for OnlyOffice:
- `guid`: Unique plugin identifier
- `EditorsSupport: ["cell"]`: Restricts plugin to spreadsheets only
- `isVisual: true`: Shows plugin panel in sidebar
- `isInsideMode: true`: Embeds UI inside OnlyOffice

### index.html

Contains the plugin UI:
- Currency dropdown (`<select id="currencySelect">`)
- Rate display box showing current exchange rate
- Three buttons: Convert to PHP, Format as Currency, Format as Accounting
- CSS styling for buttons, dropdown, and info boxes

### scripts/code.js

#### Plugin Initialization
```javascript
window.Asc.plugin.init = function() {
    fetchExchangeRates();  // Fetch rates when plugin loads
};
```

#### Exchange Rate Fetching
```javascript
function fetchExchangeRates() {
    // Calls frankfurter.app API
    xhr.open("GET", "https://api.frankfurter.app/latest?from=PHP", true);
}
```

The API returns rates as "1 PHP = X foreign currency". We invert these to get "1 foreign currency = X PHP":
```javascript
exchangeRates[currency] = 1 / data.rates[currency];
```

#### Currency Dropdown Population
```javascript
function populateCurrencyDropdown() {
    // Creates <option> elements for each currency
    // Prioritizes common currencies (USD, EUR, etc.)
    // Adds remaining currencies alphabetically
}
```

#### Rate Display Update
```javascript
function updateRateDisplay() {
    // Shows: "1 USD = ₱56.1234"
    // Updates when dropdown selection changes
}
```

#### Cell Formatting Functions
```javascript
function formatPHP() {
    oRange.SetNumberFormat("[$₱-3409]#,##0.00");
}
```
- `[$₱-3409]`: Currency symbol with Filipino locale code
- `#,##0.00`: Number format with thousands separator and 2 decimals

#### Currency Conversion
```javascript
function convertToPHP() {
    // 1. Get selected currency and exchange rate
    // 2. For each selected cell:
    //    - Read current numeric value
    //    - Multiply by exchange rate
    //    - Write new value
    //    - Apply PHP currency format
}
```

Uses `callCommand()` to execute code in OnlyOffice's context:
```javascript
window.Asc.plugin.callCommand(function() {
    oRange.ForEach(function(cell) {
        var newValue = parseFloat(value) * rate;
        cell.SetValue(newValue);
        cell.SetNumberFormat("[$₱-3409]#,##0.00");
    });
}, true, false, [{rate: rate}]);
```

The `[{rate: rate}]` parameter passes the exchange rate into the sandboxed execution context.

---

## API Reference

**Exchange Rate API**: [frankfurter.app](https://www.frankfurter.app/)
- Free, no API key required
- Updates daily on business days
- Data sourced from European Central Bank

---

## Installation

1. Open OnlyOffice Spreadsheet
2. Go to **Plugins** > **Plugin Manager**
3. Click **Install plugin manually**
4. Select `onlyofficePH-260131.plugin`

---

## Usage

1. Open a spreadsheet
2. Click **PHP Currency** in the Plugins tab
3. To convert values:
   - Select cells containing numbers
   - Choose source currency from dropdown
   - Click **Convert to PHP**
4. To format only (no conversion):
   - Select cells
   - Click **Format as ₱ Currency** or **Format as ₱ Accounting**

---

## License

MIT License - Free to use and modify.

---

*Generated on 2026-01-31 by Claude Code (Anthropic)*
