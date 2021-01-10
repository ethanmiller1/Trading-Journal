# Built-in Functions

The Trading Journal workbook has a number of functions built into it that are always available. They are listed here in alphabetical order.

| Built-in Functions                      |                                         |                                               |                                         |                                               |
| --------------------------------------- | :-------------------------------------- | :-------------------------------------------- | :-------------------------------------- | :-------------------------------------------- |
| [GetOptionType()](#GetOptionType())     | [GetNthWord()](#GetNthWord())           | [GetExpiration()](#GetExpiration())           | [GetSymbol()](#GetSymbol())             | [DaysTillExp()](#DaysTillExp())               |
| [GetStrategy()](#GetStrategy())         | [GetPosture()](#GetPosture())           | [GetStockQuote()](#GetStockQuote())           | [GetQuoteValue()](#GetQuoteValue())     | [GetPrem()](#GetPrem())                       |
| [GetMaxProfit()](#GetMaxProfit())       | [GetRisk()](#GetRisk())                 | [GetPLClose()](#GetPLClose())                 | [GetPLPercent()](#GetPLPercent())       | [GetOptionSignature()](#GetOptionSignature()) |
| [GetCommission()](#GetCommission())     | [GetOptimalDTE()](#GetOptimalDTE())     | [GetActualMaxProfit()](#GetActualMaxProfit()) | [GetPercentOfMaxProfit()](#GetPercentOfMaxProfit()) | [GetTarget1()](#GetTarget1())     |
| [GetTarget2()](#GetTarget2())           | [ExampleFunction()](#ExampleFunction()) | [ExampleFunction()](#ExampleFunction())       | [ExampleFunction()](#ExampleFunction()) | [ExampleFunction()](#ExampleFunction())       |



## GetOptionType(text)

Returns a s representing the type of option contract input by the user. Arguments may be a s copied from the thinkorswim platform.

```excel
=GetOptionType("SOLD -3 IRON CONDOR SPY 100 21 APR 17 240.5/241.5/228.5/227.5 CALL/PUT @.37")
```

Supported option strategies currently include:

| Option Type        | Example String                                                            |
| ------------------ | :------------------------------------------------------------------------ |
| Long Call          | BOT +1 FAST 100 16 FEB 18 55 CALL @1.75 LMT                               |
| Short Call         | SOLD -1 FAST 100 16 FEB 18 55 CALL @1.60 LMT                              |
| Long Put           | BOT +1 FAST 100 16 FEB 18 52.5 PUT @1.70 LMT                              |
| Short Put          | SOLD -1 FAST 100 16 FEB 18 52.5 PUT @1.55 LMT                             |
| Long Call Vertical | BOT +1 VERTICAL MRK 100 20 OCT 17 65/67.5 CALL @1.13                      |
| Short Put Vertical | SOLD -1 VERTICAL ADBE 100 19 JAN 18 170/175 CALL @2.80                    |
| Iron Condor        | SOLD -1 IRON CONDOR QCOM 100 16 FEB 18 65/67.5/60/57.5 CALL/PUT @1.44     |
| Butterfly          | BOT +1 BUTTERFLY SINA 100 19 MAY 17 65/70/75 CALL @.80                    |
| Calendar           | BOT +5 CALENDAR FSLR 100 16 JUN 17/19 MAY 17 25 CALL @.31                 |
| Diagonal           | BOT +1 DIAGONAL CRM 100 18 AUG 17/21 JUL 17 87.5/82.5 PUT @2.57           |
| Synthetic          | SOLD -1 COMBO CVX 100 18 AUG 17 105 CALL/PUT @-2.06 LMT [TO OPEN/TO OPEN] |



## GetNthWord(text, start_num[, num_words])

Returns the nth word in a string. It takes in a string for the first argument, and an integer representing the nth word you would like returned as the second argument. An optional argument is an integer representing the number of words you want returned after the nth word (the default value is `1`).

To parse the date from a Vertical TOS data string, for example, employ the following usage:

```excel
=GetNthWord("BOT +1 VERTICAL MRK 100 20 OCT 17 65/67.5 CALL @1.13", 6, 3)
```

This returns `20 OCT 17`.



## GetExpiration(trade_order, option_type)

Returns the expiration date from a TOS data string passed in by the user. Arguments may be a TOS data string and a string representing the option type being evaluated. The following usage would return `10/20/2017`.

```excel
=GetExpiration("BOT +1 VERTICAL MRK 100 20 OCT 17 65/67.5 CALL @1.13","Vertical")
```



## GetSymbol(trade_order, option_type)

Returns a ticker symbol. The first argument is a TOS data string and the second is a string representing the option type being evaluated. The following usage would return `FAST`.

```excel
=GetSymbol("BOT +1 FAST 100 16 FEB 18 55 PUT @1.75 LMT", "Put")
```



## DaysTillExp(trade_date, expiration_date)

Returns an integer representing the number of days the order expiration was removed from the date your TOS order was filled. Both arguments are a date. The following usage would return `28`.

```excel
=DaysTillExp("1/19/2018", "2/16/2018")
```



## GetStrategy(trade_order, option_type)

Returns a string representing the option strategy of a TOS order. The first argument is a TOS order and the second is a string representing the option type being evaluated. The following usage would return `Long Put Diagonal`.

```excel
=GetStrategy("BOT +1 DIAGONAL CRM 100 18 AUG 17/21 JUL 17 87.5/82.5 PUT @2.57", "Diagonal")
```



## GetPosture(option_strategy)

Returns a string representing the market posture of a TOS order (bullish, bearish, or neutral). Arguments may be a string returned by the GetStrategy() function. The following usage would return `Bearish`.

```excel
=GetPosture("Long Put")
```



## GetStockQuote(symobol, date)

Returns a JSON stock quote. The first argument is a string representing a company's ticker symbol, and the second argument is the date you want to query. The following usage would return `{"date":"2018-01-19","uClose":54.06,"uOpen":53.7,"uHigh":54,"uLow":54.09,"uVolume":3605678,"close":27.07,"open":27.76,"high":27.11,"low":27.14,"volume":7067646,"change":0.2368,"changePercent":0.8807,"label":"Jan 19, 18","changeOverTime":0.26904}`.

```excel
=GetStockQuote("FAST","1/19/2018")
```

Note: this function is currently set to query stock quotes up to 2 years in the past. This can be changed by altering the Url variable to use a different element in the ranges array. Press `Alt+F11` to open VBA, select module1, and look for the `GetStockQuote()` function. Possible arguments include:

```vb
ranges = Array("5d", "1m", "3m", "6m", "1y", "2y", "5y", "max")
```

This function uses [iexcloud's API](https://iexcloud.io/console/search) to pull stock data. Alternative API's are available to use:

1.  [iexcloud](https://iexcloud.io/console/search)
2.  [Intrinio](https://intrinio.com/)
3.  [TD Ameritrade](https://www.reddit.com/r/algotrading/comments/914q22/successful_access_to_td_ameritrade_api/)

Be sure to review the [documentation](https://iexcloud.io/docs/api/#historical-prices) for iex cloud's historical prices api.



## GetQuoteValue(key, stock_quote)

Returns the value associated with a key from a JSON. The first argument is a string representing the desired key, and the second argument is the stock quote JSON. The following usage would return `27.22`.

```excel
=GetQuoteValue("high","{"date":"2018-01-19","uClose":56.29,"uOpen":54.05,"uHigh":55,"uLow":54.57,"uVolume":3537791,"close":27.03,"open":27.16,"high":27.22,"low":26.98,"volume":7277005,"change":0.2314,"changePercent":0.8633,"label":"Jan 19, 18","changeOverTime":0.26099}")
```

Valid keys are:

| Keys           |
| -------------- |
| date           |
| uClose         |
| uOpen          |
| uHigh          |
| uLow           |
| uVolume        |
| close          |
| open           |
| high           |
| low            |
| volume         |
| change         |
| changePercent  |
| label          |
| changeOverTime |



## GetPrem(trade_order, option_type)

Returns a double representing the option premium of a TOS order. The first argument is a TOS order and the second is a string representing the option type being evaluated. The following usage would return `1.75`.

```excel
=GetPrem("BOT +1 FAST 100 16 FEB 18 55 CALL @1.75 LMT", "Call")
```



## GetMaxProfit(trade_order, option_type, qty, prem[, comm][,risk])

Returns the max profit calculated from a TOS order. The first argument is a TOS order, the second argument is a string representing the option type of the TOS order, the third argument is the number of contracts, the fourth argument is the premium per share, the fifth argument is commissions paid to your broker, and is the sixth argument is the risk of the trade. The following usage would return `137`.

```excel
=GetMaxProfit("BOT +1 VERTICAL MRK 100 20 OCT 17 65/67.5 CALL @1.13
","Vertical", 1, 1.13)
```

Note: Risk only needs to be included to estimate max profits for calendar spreads.

The following is a list of formulas included in this function for calculating the max profit of currently supported strategies:

| Strategy                                                                 | Max Profit                                                                         | Formula                                    |
| ------------------------------------------------------------------------ | ---------------------------------------------------------------------------------- | ------------------------------------------ |
| [Iron Condor](https://www.tastytrade.com/tt/learn/iron-condor)           | net credit received                                                                | Premium                                    |
| [Butterfly Spread](https://www.tastytrade.com/tt/learn/butterfly-spread) | distance between the short strike and long strike, less the debit paid             | \| Short Strike - Long Strike \| - Premium |
| [Calendar Spread](https://www.tastytrade.com/tt/learn/calendar-spread)   | incalculable: occurs if stock price = strike price at front-month expiration\*     | ~Short Call Credit - Net Loss on Long Call |
| [Diagonal Spread](https://www.tastytrade.com/tt/learn/diagonal-spread)   | incalculable due to the differing expiration cycles                                | ~Spread Width - Premium + Extrinsic Value  |
| [Synthetics](https://www.tastytrade.com/tt/learn/synthetics)             | undefined (unlimited)                                                              | Strike + Net Credit Received               |
| [Vertical Spread](https://www.tastytrade.com/tt/learn/vertical-spread)   | distance between strikes less net debit paid OR credit received from opening trade | Spread Width - Premium OR Premium          |
| [Short Call/Put](https://www.tastytrade.com/tt/learn/naked-options)      | credit received from opening trade                                                 | Premium                                    |
| [Long Call](#GetMaxProfit)                                               | undefined (unlimited)                                                              | ∞                                          |
| [Long Put](#GetMaxProfit)                                                | strike price less debit paid from opening trade                                    | Strike - Premium                           |

Note: all euqations must be multiplied by the number of shares being controlled (# of contracts \* 100) and subtract commissions.

\*Formula must use the Black-Scholes model to calculate the theoretical value of the Long Call when the Short Call is worthless. You make money on the short call and lose money on the long call. The key to max profit is making as much money on the short call as you can, and losing as little money on the long call as you can.



## GetRisk(trade_order, option_type, qty, prem, max_profit[, comm])

Returns the risk calculated from a TOS order. The first argument is a TOS order, the second argument is a string representing the option type of the TOS order, the third argument is the number of contracts, the fourth argument is the premium per share, the fifth argument is the max profit calculated from a TOS order, and the sixth argument is commissions paid to your broker. The following usage would return `113`.

```excel
=GetRisk("BOT +1 VERTICAL MRK 100 20 OCT 17 65/67.5 CALL @1.13
","Vertical", 1, 1.13, 137)
```

The following is a list of formulas included in this function for calculating the risk of currently supported strategies:

| Strategy                                                               | Max Profit                                                                        | Formula                           |
| ---------------------------------------------------------------------- | --------------------------------------------------------------------------------- | --------------------------------- |
| [Iron Condor](https://www.tastytrade.com/tt/learn/iron-condor)         | the greater of the two vertical spreads                                           | Spread Width                      |
| [Vertical Spread](https://www.tastytrade.com/tt/learn/vertical-spread) | distance between strikes less net credit received OR debit paid for opening trade | Spread Width - Premium OR Premium |
| [Diagonal Spread](https://www.tastytrade.com/tt/learn/diagonal-spread) | distance between strikes less net credit received OR debit paid for opening trade | Spread Width - Premium OR Premium |
| [Synthetics](https://www.tastytrade.com/tt/learn/synthetics)           | undefined (unlimited)                                                             | ∞ OR Strike - Premium             |
| [Short Call](https://www.tastytrade.com/tt/learn/naked-options)        | undefined (unlimited)                                                             | ∞                                 |
| [Short Put](https://www.tastytrade.com/tt/learn/naked-options)         | strike price less credit received from opening trade                              | Strike - Premium                  |
| [Standard](#GetRisk())                                                 | debit paid for opening trade                                                      | Premium                           |

Note: all euqations must be multiplied by the number of shares being controlled (# of contracts \* 100) and add commissions.



## GetPLClose(trade_order, option_type, prem, max_profit[, comm])

Returns the profit (or loss) of a closed TOS order. The first argument is a TOS order, the second arguments is a string representing the type of option being evaluated, the third argument is a string representation of the closing option premium, and the fourth argument is the max profit of a TOS order. The following usage would return `25`.

```excel
=GetPLClose("BOT +1 FAST 100 16 FEB 18 55 PUT @1.75 LMT","Put","2.00","5,325")
```

Numeric values are passed in as strings to avoid datatype errors when cells contain null values. They are converted to the appropriate datatypes *inside* the function.

Many of the strategies in this function use max profit to calculate P/L Closed. As a consequence, commissions are accounted for. It only needs to be passed in for synthetics.



## GetPLPercent(pl_closed, max_profit, max_risk)

Returns the profit (or loss) of a closed TOS order. The first argument is the Profit/Loss dollar amount, the second arguments is the max profit, the third argument is the total risk. The following usage would return `18%`.

```excel
=GetPLPercent("43", "243", "257")
```

Numeric values are passed in as strings to avoid datatype errors when cells contain null values. They are converted to the appropriate datatypes *inside* the function.



## GetOptionSignature(trade_order)

Returns the option signature of a TOS order. It can be used to chart the price of an option over time.

```excel
=GetOptionSignature("BOT +1 VERTICAL MRK 100 20 OCT 17 65/67.5 CALL @1.13")
```

### Test Table:

| TOS Order                                                            | Option Signature                                            |
| -------------------------------------------------------------------- | :---------------------------------------------------------- |
| SOLD -1 IRON CONDOR MMM 100 20 OCT 17 200/220/195/190 CALL/PUT @8.53 | .MMM171020C200-.MMM171020C220+.MMM171020P195-.MMM171020P190 |
| BOT +1 VERTICAL MRK 100 20 OCT 17 65/67.5 CALL @1.13                 | .MRK171020C65-.MRK171020C67.5'                              |
| BOT +1 DIAGONAL CRM 100 18 AUG 17/21 JUL 17 87/82.5 PUT @2.57        | .CRM170818P87-.CRM170721P.5                                 |
| BOT +1 FAST 100 16 FEB 18 55 CALL @1.75 LMT                          | .FAST180216C55                                              |
| BOT +1 BUTTERFLY SINA 100 19 MAY 17 65/70/75 CALL @.80               | .SINA170519C65-.SINA170519C70-.SINA170519C70+.SINA170519C75 |
| SOLD -2 COMBO HPQ 100 20 OCT 17 19 CALL/PUT @.13                     | .HPQ171020C19-.HPQ171020P19                                 |
| BOT +5 CALENDAR FSLR 100 16 JUN 17/19 MAY 17 25 CALL @.31            | .FSLR170616C25-.FSLR170519C25                               |



## GetCommission(trade_order)

Returns the commission paid to thinkorswim for a fulfilled order. It takes a thinkorswim order to determine how many contracts are being evaluated. The following usage would return `$8.50`.

```excel
=GetCommission("SOLD -1 IRON CONDOR QCOM 100 16 FEB 18 65/67.5/60/57.5 CALL/PUT @1.44")
```

The formula is as follows:

```msgraph-interactive
4 contracts * $0.75 fee per contract = $3 + $1.25 base fee = $4.25 * 2 for opening and closing = $8.50
```



## GetOptimalDTE(trade_order, expiration_date)

Returns an integer representing the number of days the order expiration was removed from the date where option premium was optimal for maximum profit. The first argument is a TOS order, the second argument is a date. The following usage would return `48`.

```excel
=GetOptimalDTE("BOT +1 FAST 100 16 FEB 18 55 PUT @1.75 LMT", "12/30/2017")
```



## GetActualMaxProfit(trade_order, optimal_prem)

Returns a currency representing the highest dollar amount to be received if the trade was closed at the optimal exit date. The optimal exit date is determined as the result of an historical analysis of the underlying's price levels while the trade was open. Assuming a bullish posture, the hishest price level reached is the "actual max profit" we could have received based on real hstorical price movement. The first argument is a TOS order, the second argument is a string representing the max premium that could have been received since the trade was opened. The following usage would return `$61`.

```excel
=GetActualMaxProfit("BOT +1 FAST 100 16 FEB 18 55 PUT @1.75 LMT", 2.40)
```



## GetPercentOfMaxProfit(max_profit, risk, actual_max_profit)

Returns the percentage of the theoretical max profit that could have been received if the trade was closed at the optimal time. These data are collected to indicate what percent of max profit is generally good to close the trade at. If there was no profitable point during the lifetime of the trade, the formula will calculate the percent of max loss. All arguments are strings. The first argument is a max profit, the second argument is the risk, and the third argument is the highest amount that could have been received if the trade was closed at the optimal exit time. The following usage would return `70%`.

```excel
=GetPercentOfMaxProfit("116", "135", "81.50")
```



## GetTarget1(support, resistance, entry)

Returns the T1 technical target of a flag using the flag pole method. The flag pole method adds the length of the flag pole to the current high of the low day (which is also the entry). T1 occurs at 161.8% of the projected second flag pole. The following usage would return `33.01`.

```excel
=GetTarget1("30.36", "32.35", "31.78")
```

The purpose tracking both T1 and T2 targets of each order is to determine how often T2 is reached compared to T1 and isolate the conditions surrounding the observed performance. Condition markers may be:

1. Ticker Symbol
1. Trend
1. Posture
1. Price Pattern

This data will be considered in crafting target rules.

## Column Explainations:

| Column Alias | Column Name | Example String                                                           |
| ------------ | :---------- |:------------------------------------------------------------------------ |
| [prot]       | Protection  | The next low (i.e. a candle that moves lower than yesterday's low). If we have a series to higher lows, and then a candle has a low that's lower than yesterday's low, we want to add a stop 20 cents below the new low in case it continues dropping. What if it closes below yesterday's close? |
| Last R (R3)   | Last Resistance  | New Resistance. If bullish, the next high. If bearish, the next low. New high (or low) becomes new resistance. The highest price level the underlying reached while our trade was open. This data is used to determine target rules based on how close to T1 and T2 targerts the underlying reached. |
| Last S (S3)   | Last Support  | New Support. If bullish, the next low. If bearish, the next high. New low (or high) becomes new support. The lowest price level the underlying reached while our trade was open. This data is used to determine what appropriate stop rules are. |

#### 3 Rule Outcomes

1. Saved - Would a stop rule have saved us from losing a lot of money before it reached an extreme low?
1. Damaged - Would a stop rule have caused us to exit the trade right before the stock went to our target?
1. Unaffected - Would a stop rule have been irrelevent because the lowest price never reached it anyway?

#### Stop Rules

1. R1 % Stop Loss - Resistance minus 1% of resistance

![](https://i.ibb.co/ykHV8WF/image.png) 

2. R1 Fixed Stop Loss - Resistance minus 20 cents
1. S2 % Stop Loss - Support minus 1%
1. S2 Fixed Stop Loss - Support minus 20 cents      
1. P1 %
1. P1 Fixed
1. Max Loss % - 50% of an option's max loss

## GetTarget2(support, resistance, entry)

Returns the T2 technical target of a flag using the flag pole method. The flag pole method adds the length of the flag pole to the current high of the low day (which is also the entry). T2 occurs at 200% of the projected second flag pole. The following usage would return `33.77`.

```excel
=GetTarget2("30.36", "32.35", "31.78")
```


## GetStopLossRule()