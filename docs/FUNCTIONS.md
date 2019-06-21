# Built-in Functions

The Trading Journal workbook has a number of functions built into it that are always available. They are listed here in alphabetical order.

| Built-in Functions                  |                    |       |       |       |
| ------------------------------------|:-------------------| :-----| :-----| :-----|
| [getOptionType()](#getOptionType()) | [getNthWord()](#getNthWord()) | [getExpiration()](#getExpiration()) | [getSymbol()](#getSymbol()) | [daysTillExp()](#daysTillExp()) |
| [getStrategy()](#getStrategy())  | [getPosture()](#getPosture()) | [getStockQuote()](#getStockQuote()) | [exampleFunction()](#exampleFunction()) | [exampleFunction()](#exampleFunction()) |
| [exampleFunction()](#exampleFunction()) | [exampleFunction()](#exampleFunction()) | [exampleFunction()](#exampleFunction()) | [exampleFunction()](#exampleFunction()) | [exampleFunction()](#exampleFunction()) |

<a name="getOptionType()"></a>
## getOptionType(text)
Returns a string representing the type of option contract input by the user. Arguments may be a string copied from the thinkorswim platform.

``` excel
=getOptionType("SOLD -3 IRON CONDOR SPY 100 21 APR 17 240.5/241.5/228.5/227.5 CALL/PUT @.37")
```

Supported option strategies currently include:

| Option Type       | Example String |
|-------------------|:--------------|
|Long Call          |BOT +1 FAST 100 16 FEB 18 55 CALL @1.75 LMT|
|Short Call         |SOLD -1 FAST 100 16 FEB 18 55 CALL @1.60 LMT|
|Long Put           |BOT +1 FAST 100 16 FEB 18 52.5 PUT @1.70 LMT|
|Short Put          |SOLD -1 FAST 100 16 FEB 18 52.5 PUT @1.55 LMT|
|Long Call Vertical |BOT +1 VERTICAL MRK 100 20 OCT 17 65/67.5 CALL @1.13|
|Short Put Vertical |SOLD -1 VERTICAL ADBE 100 19 JAN 18 170/175 CALL @2.80|
|Iron Condor        |SOLD -1 IRON CONDOR QCOM 100 16 FEB 18 65/67.5/60/57.5 CALL/PUT @1.44|
|Butterfly          |BOT +1 BUTTERFLY SINA 100 19 MAY 17 65/70/75 CALL @.80|
|Calendar           |BOT +5 CALENDAR FSLR 100 16 JUN 17/19 MAY 17 25 CALL @.31|
|Diagonal           |BOT +1 DIAGONAL CRM 100 18 AUG 17/21 JUL 17 87.5/82.5 PUT @2.57|
|Synthetic          |SOLD -1 COMBO CVX 100 18 AUG 17 105 CALL/PUT @-2.06 LMT [TO OPEN/TO OPEN]|

<a name="getNthWord()"></a>
## getNthWord(text, start_num[, num_words])

Returns the nth word in a string. It takes in a string for the first argument, and an integer representing the nth word you would like returned as the second argument. An optional argument is an integer representing the number of words you want returned after the nth word (the default value is `1`).

To parse the date from a Vertical TOS data string, for example, employ the following usage:

``` excel
=getNthWord("BOT +1 VERTICAL MRK 100 20 OCT 17 65/67.5 CALL @1.13", 6, 3)
```
This returns `20 OCT 17`.

<a name="getExpiration()"></a>
## getExpiration(trade_order, option_type)

Returns the expiration date from a TOS data string passed in by the user. Arguments may be a TOS data string and a string representing the option type being evaluated. The following usage would return `10/20/2017`.

``` excel
=getExpiration("BOT +1 VERTICAL MRK 100 20 OCT 17 65/67.5 CALL @1.13","Vertical")
```

<a name="getSymbol()"></a>
## getSymbol(trade_order, option_type)

Returns a ticker symbol. The first argument is a TOS data string and the second is a string representing the option type being evaluated. The following usage would return `FAST`.

``` excel
=getSymbol("BOT +1 FAST 100 16 FEB 18 55 PUT @1.75 LMT", "Put")
```

<a name="daysTillExp()"></a>
## daysTillExp(trade_date, expiration_date)

Returns an integer representing the number of days the order expiration was removed from the date your TOS order was filled. Both arguments are a date. The following usage would return `28`.

``` excel
=daysTillExp("1/19/2018", "2/16/2018")
```

<a name="getStrategy()"></a>
## getStrategy(trade_order, option_type)

Returns a string representing the option strategy of a TOS order. The first argument is a TOS order and the second is a string representing the option type being evaluated. The following usage would return `Long Put Diagonal`.

``` excel
=getStrategy("BOT +1 DIAGONAL CRM 100 18 AUG 17/21 JUL 17 87.5/82.5 PUT @2.57", "Diagonal")
```

<a name="getPosture()"></a>
## getPosture(option_strategy)

Returns a string representing the market posture of a TOS order (bullish, bearish, or neutral). Arguments may be a string returned by the getStrategy() function. The following usage would return `Bearish`.

``` excel
=getPosture("Long Put")
```

<a name="getStockQuote()"></a>
## getStockQuote(option_strategy)

Returns a JSON stock quote. The first argument is a string representing a companies ticker symbol, and the second argument is the date you want to query. The following usage would return `{"date":"2018-01-19","uClose":54.06,"uOpen":53.7,"uHigh":54,"uLow":54.09,"uVolume":3605678,"close":27.07,"open":27.76,"high":27.11,"low":27.14,"volume":7067646,"change":0.2368,"changePercent":0.8807,"label":"Jan 19, 18","changeOverTime":0.26904}`.

``` excel
=getStockQuote("FAST","1/19/2018")
```

Note: this function is currently set to query stock quotes up to 2 years in the past. This can be changed by altering the Url variable to use a different element in the ranges array. Press `Alt+F11` to open VBA, select module1, and look for the `getStockQuote()` function. Possible arguments include:

``` vb
ranges = Array("5d", "1m", "3m", "6m", "1y", "2y", "5y", "max")
```

This function uses [iexcloud's API](https://iexcloud.io/console/search) to pull stock data. Alternative API's are available to use:

1. [iexcloud](https://iexcloud.io/console/search)
1. [Intrinio](https://intrinio.com/)
1. [TD Ameritrade](https://www.reddit.com/r/algotrading/comments/914q22/successful_access_to_td_ameritrade_api/)

Be sure to review the [documentation](https://iexcloud.io/docs/api/#historical-prices) for iex cloud's historical prices api.