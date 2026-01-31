(function(window, undefined) {

    var exchangeRates = {};
    var lastUpdated = null;

    window.Asc.plugin.init = function() {
        console.log("PHP Currency plugin initialized");
        fetchExchangeRates();
    };

    window.Asc.plugin.button = function(id) {
        this.executeCommand("close", "");
    };

    // Fetch exchange rates from API
    function fetchExchangeRates() {
        var xhr = new XMLHttpRequest();
        // Using frankfurter.app - free, no API key needed
        xhr.open("GET", "https://api.frankfurter.app/latest?to=PHP", true);

        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        var data = JSON.parse(xhr.responseText);
                        processRates(data);
                    } catch (e) {
                        showError("Error parsing exchange rate data");
                    }
                } else {
                    showError("Could not fetch exchange rates");
                }
            }
        };

        xhr.onerror = function() {
            showError("Network error fetching rates");
        };

        xhr.send();
    }

    function processRates(data) {
        // frankfurter returns rates FROM base currency TO PHP
        // We need to store how much 1 unit of each currency = X PHP
        var phpRate = data.rates.PHP;
        var baseCurrency = data.base; // EUR

        // Calculate rates for common currencies to PHP
        // First, get rates from EUR to other currencies, then to PHP
        fetchAllRates();
    }

    function fetchAllRates() {
        var xhr = new XMLHttpRequest();
        xhr.open("GET", "https://api.frankfurter.app/latest?from=PHP", true);

        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        var data = JSON.parse(xhr.responseText);
                        // data.rates contains how many units of each currency = 1 PHP
                        // We need the inverse: how many PHP = 1 unit of currency
                        exchangeRates = {};
                        lastUpdated = data.date;

                        for (var currency in data.rates) {
                            exchangeRates[currency] = 1 / data.rates[currency];
                        }

                        populateCurrencyDropdown();
                        updateRateDisplay();
                    } catch (e) {
                        showError("Error parsing exchange rate data");
                    }
                } else {
                    showError("Could not fetch exchange rates");
                }
            }
        };

        xhr.onerror = function() {
            showError("Network error fetching rates");
        };

        xhr.send();
    }

    function populateCurrencyDropdown() {
        var select = document.getElementById("currencySelect");
        select.innerHTML = "";

        var currencies = {
            "USD": "US Dollar",
            "EUR": "Euro",
            "GBP": "British Pound",
            "JPY": "Japanese Yen",
            "AUD": "Australian Dollar",
            "CAD": "Canadian Dollar",
            "CHF": "Swiss Franc",
            "CNY": "Chinese Yuan",
            "HKD": "Hong Kong Dollar",
            "SGD": "Singapore Dollar",
            "KRW": "South Korean Won",
            "THB": "Thai Baht",
            "MYR": "Malaysian Ringgit",
            "IDR": "Indonesian Rupiah",
            "INR": "Indian Rupee",
            "NZD": "New Zealand Dollar",
            "MXN": "Mexican Peso",
            "BRL": "Brazilian Real",
            "AED": "UAE Dirham",
            "SAR": "Saudi Riyal"
        };

        for (var code in currencies) {
            if (exchangeRates[code]) {
                var option = document.createElement("option");
                option.value = code;
                option.textContent = code + " - " + currencies[code];
                select.appendChild(option);
            }
        }

        // Add any other available currencies
        for (var code in exchangeRates) {
            if (!currencies[code]) {
                var option = document.createElement("option");
                option.value = code;
                option.textContent = code;
                select.appendChild(option);
            }
        }

        select.onchange = updateRateDisplay;
        document.getElementById("convertBtn").disabled = false;
    }

    function updateRateDisplay() {
        var select = document.getElementById("currencySelect");
        var rateInfo = document.getElementById("rateInfo");
        var currency = select.value;

        if (currency && exchangeRates[currency]) {
            var rate = exchangeRates[currency].toFixed(4);
            rateInfo.className = "rate-info";
            rateInfo.innerHTML = "1 " + currency + " = <span class='rate-value'>₱" + formatNumber(exchangeRates[currency]) + "</span>" +
                "<div class='last-updated'>Last updated: " + lastUpdated + "</div>";
        }
    }

    function formatNumber(num) {
        return num.toLocaleString('en-PH', {minimumFractionDigits: 2, maximumFractionDigits: 4});
    }

    function showError(message) {
        var rateInfo = document.getElementById("rateInfo");
        rateInfo.className = "rate-info error";
        rateInfo.textContent = message;
    }

    // Make functions available globally
    window.exchangeRates = exchangeRates;
    window.getSelectedCurrency = function() {
        return document.getElementById("currencySelect").value;
    };
    window.getExchangeRate = function(currency) {
        return exchangeRates[currency];
    };

})(window);

// Format selected cells as PHP currency
function formatPHP() {
    window.Asc.plugin.callCommand(function() {
        var oRange = Api.GetSelection();
        if (oRange) {
            oRange.SetNumberFormat("[$₱-3409]#,##0.00");
        }
    }, true);
}

// Format selected cells as PHP accounting
function formatPHPAccounting() {
    window.Asc.plugin.callCommand(function() {
        var oRange = Api.GetSelection();
        if (oRange) {
            oRange.SetNumberFormat("[$₱-3409]#,##0.00;([$₱-3409]#,##0.00)");
        }
    }, true);
}

// Convert selected cells from chosen currency to PHP
function convertToPHP() {
    var currency = document.getElementById("currencySelect").value;
    var rate = window.getExchangeRate(currency);

    if (!currency || !rate) {
        alert("Please select a currency first");
        return;
    }

    window.Asc.plugin.callCommand(function() {
        var rate = window.Asc.plugin.info.rate;
        var oSheet = Api.GetActiveSheet();
        var oRange = Api.GetSelection();

        if (oRange) {
            var address = oRange.GetAddress(true, true, "xlA1", false);
            // Parse the address to get cell references
            var cells = [];

            // Handle range like "A1:B5" or single cell "A1"
            oRange.ForEach(function(cell) {
                var value = cell.GetValue();
                if (value && !isNaN(parseFloat(value))) {
                    var newValue = parseFloat(value) * rate;
                    cell.SetValue(newValue);
                    cell.SetNumberFormat("[$₱-3409]#,##0.00");
                }
            });
        }
    }, true, false, [{rate: rate}]);
}
