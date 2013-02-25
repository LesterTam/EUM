Attribute VB_Name = "FINAN_ASSET_FUTURES_FAIR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
'The most widely traded equity index futures contract in the U.S. is the S&P 500.
'The futures contracts on the S&P 500 index are traded at the Chicago Mercantile
'Exchange (CME). The value of the contract is $250 times the futures price. The CME-Miniù
'contract is a smaller, electronically-traded version of the original pit-traded contract
'and has a value of $50 times the futures price. So, if the futures contract
'was valued at 1000, it would have a notional value of $250,000 and the CME-Mini a
'notional value of $50,000. The CME also trades options on these futures contracts. The
'Chicago Board Options Exchange(CBOE) trades options on the cash S&P 500 index. The S&P 500
'Index consists of 500 stocks, each selected for their market size, liquidity, and industry
'group. Also, the S&P 500 is a market value weighted index where the market value of an
'individual stock is the stock price times the number of shares outstanding. Each stock's
'weight in the Index then is proportionate to its market value. The weights for the
'individual stocks change as their respective prices rise and fall relative to other stocks in
'the index (Kolb,1997). Alternatively, an index could be price weighted, where the index weights
'are proportional to the stock prices. The Dow Jones Industrial Average is an example of a
'price weighted index.
'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------

Function FUTURES_INDEX_FAIR_PRICE_FUNC( _
ByVal MATURITY As Date, _
Optional ByVal INDEX_STR As String = "^DJI", _
Optional ByVal RATE_STR As String = "^TNX", _
Optional ByVal REFRESH_CALLER As Variant, _
Optional ByVal SERVER_STR As String = "UNITED STATES", _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal VERSION As Integer = 0)

'^TNX: 10-YEAR TREASURY NOTE(Chicago Options: ^TNX)

Dim i As Long
Dim j As Long 'Counter for components
Dim NROWS As Long

Dim SETTLEMENT As Date

Dim TEMP_SUM As Double 'SUM of Prices (components/stocks)
Dim DIVISOR_VAL As Double
'A $1 change in any stock price will change the Index by 1 / DIVISOR_VAL
Dim YEARFRAC_VAL As Double 'Years to Expiry
Dim AVG_YIELD_VAL As Double
Dim DIVIDENDS_VAL As Double 'Total Dividends for X days
Dim RISK_FREE_RATE_VAL As Double 'Annual Risk-free Rate
Dim INDEX_QUOTE_VAL As Double 'Current Value of the Index
Dim INDEX_FAIR_VAL As Double
Dim SPREAD_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

SETTLEMENT = Now
YEARFRAC_VAL = YEARFRAC_FUNC(SETTLEMENT, MATURITY, COUNT_BASIS)
TEMP_MATRIX = MATRIX_YAHOO_QUOTES_FUNC(RATE_STR & "," & INDEX_STR, "l1", SERVER_STR, REFRESH_CALLER, False, "+")
RISK_FREE_RATE_VAL = TEMP_MATRIX(1, 1)
RISK_FREE_RATE_VAL = RISK_FREE_RATE_VAL / 100
INDEX_QUOTE_VAL = TEMP_MATRIX(2, 1)


ReDim TEMP_MATRIX(1 To 1, 1 To 4)
TEMP_MATRIX(1, 1) = "Name"
TEMP_MATRIX(1, 2) = "Symbol"
TEMP_MATRIX(1, 3) = "Last Trade"
TEMP_MATRIX(1, 4) = "dividend yield"

DATA_MATRIX = YAHOO_INDEX_QUOTES_FUNC(INDEX_STR, TEMP_MATRIX, "", _
              True, REFRESH_CALLER, SERVER_STR)
NROWS = UBound(DATA_MATRIX, 1)

j = 0
TEMP_SUM = 0
AVG_YIELD_VAL = 0
For i = 1 To NROWS
    TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, 3)
'-------------------------------------------------------------------------------
    If VERSION = 0 Then 'Exclude non-dividend components
'-------------------------------------------------------------------------------
        If DATA_MATRIX(i, 4) <> 0 Then
            AVG_YIELD_VAL = AVG_YIELD_VAL + DATA_MATRIX(i, 4) / 100
            j = j + 1
        End If
'-------------------------------------------------------------------------------
    Else
'-------------------------------------------------------------------------------
        AVG_YIELD_VAL = AVG_YIELD_VAL + DATA_MATRIX(i, 4) / 100
        j = j + 1
'-------------------------------------------------------------------------------
    End If
'-------------------------------------------------------------------------------
Next i

DIVISOR_VAL = TEMP_SUM / INDEX_QUOTE_VAL
AVG_YIELD_VAL = AVG_YIELD_VAL / j
DIVIDENDS_VAL = TEMP_SUM * AVG_YIELD_VAL * YEARFRAC_VAL
INDEX_FAIR_VAL = INDEX_QUOTE_VAL * (1 + RISK_FREE_RATE_VAL * YEARFRAC_VAL) - DIVIDENDS_VAL
SPREAD_VAL = INDEX_FAIR_VAL - INDEX_QUOTE_VAL

ReDim TEMP_MATRIX(1 To 12, 1 To 2)

TEMP_MATRIX(1, 1) = "INDEX"
TEMP_MATRIX(1, 2) = INDEX_STR

TEMP_MATRIX(2, 1) = "RISK FREE RATE"
TEMP_MATRIX(2, 2) = RATE_STR

TEMP_MATRIX(3, 1) = "INDEX FAIR VALUE"
TEMP_MATRIX(3, 2) = INDEX_FAIR_VAL

TEMP_MATRIX(4, 1) = "RISK FREE RATE"
TEMP_MATRIX(4, 2) = RISK_FREE_RATE_VAL

TEMP_MATRIX(5, 1) = "INDEX QUOTE"
TEMP_MATRIX(5, 2) = INDEX_QUOTE_VAL

TEMP_MATRIX(6, 1) = "SPREAD"
TEMP_MATRIX(6, 2) = SPREAD_VAL

TEMP_MATRIX(7, 1) = "DIVIDENDS VALUE"
TEMP_MATRIX(7, 2) = DIVIDENDS_VAL

TEMP_MATRIX(8, 1) = "AVERAGE YIELD"
TEMP_MATRIX(8, 2) = AVG_YIELD_VAL

TEMP_MATRIX(9, 1) = "DIVISOR VALUE"
TEMP_MATRIX(9, 2) = DIVISOR_VAL

TEMP_MATRIX(10, 1) = "SUM COMPONENTS PRICES"
TEMP_MATRIX(10, 2) = TEMP_SUM

TEMP_MATRIX(11, 1) = "NO COMPONENTS"
TEMP_MATRIX(11, 2) = j

TEMP_MATRIX(12, 1) = "TENOR"
TEMP_MATRIX(12, 2) = YEARFRAC_VAL

FUTURES_INDEX_FAIR_PRICE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
FUTURES_INDEX_FAIR_PRICE_FUNC = Err.number
End Function

'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
'DOW futures
'REFERENCES:
'http://www.cbot.com/cbot/pub/page/0,3181,1165,00.html
'http://www.cbot.com/cbot/pub/cont_detail/1,3206,1719+8708,00.html
'http://www.indexarb.com/dividendYieldSorteddj.html
'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
'Every morning, before the local stock market opens, I check the futures, in particular,
'the DOW futures. Suppose the current DOW index is at 8750. You buy one DOW future and it'll
'cost you $87,500 that's 10 times the current value of the DOW index. The futures contract
'has an expiry date. Whatever the DOW index is, at that date, that's what you'll get for
'your future contract, multiplied by 10. If the DOW is at 8635, you'll get $86,350, and you
'lose a bundle.If the DOW is 8942, you'll get $89,420, and you make a bundle.

'These futures contracts are another way of trading the DOW. They're like Options for the DOW.
'You can also always buy a mutual fund that tracks the DOW, or you can buy all 30 stocks in
'the DOW. They are offered by the Chicago Board of Trade.

'And the multiplier is always 10 for the DOW. It's 250 for the S&P and 100 for Nasdaq. There
'are also mini-DOW futures where the multiplier is 5. However, unlike Options (which give the
'purchaser the right, but not the obligation, to buy), the buyer is obligated to buy.

'At Expiry, a Special Opening Quotation is calculated and money (not stocks!) changes hands.
'But them futures are worth a fortune! Actually, you can buy a DOW future with just, say, $6750.
'That's a HUGE leverage. It 's about 4% to 7% of the value of the contract. If you were to buy
'all 30 stocks on the DOW, you'd have to put up (about) 50% of the price.

'At the end of each day, gains (or losses) are credited to futures accounts. If your account
'balance drops below the margin, you'll have to add $$$. So you can make a fortune, or lose a
'fortune.

'However, since DOW futures start trading about an hour before the market opens, I can just
'take a peek to get an idea of how the DOW will move, when the market DOES open.

'DOW futures are an alternate way to invest in the 30 stocks of the DOW. However, if you had
'bought all 30 stocks in the DOW, you'd get all the dividends. If you buy DOW futures, you get
'NO dividends. Hence the "price" of the future is reduced to account for the series of dividends
'... until the futures contract expires.

'Above, we said the futures contract was 10 times the current value of the DOW index.
'Actually, if the DOW Futures are trading at, say, 8750, you'd pay 10x8750 for the contract.
'In other words, it's 10 times the current value of the DOW Futures ... not 10 times the
'value of the DOW Index.

'At one time, the DOW was the simple, garden variety average of 30 stock prices:
'(P1 + P2 + ... + P30)/30. But, over time, the set of 30 changed (some were replaced
'by others) and there were stock splits (so prices doubled or tripled) etc.
'so the number 30 got changed so that the Index wouldn't have a discontinuity.

'The magic divisor is now quite different: As I write this, the sum of prices is $1171.08
'and the DOW Index is 9325.01 so 9325.01 = 1171.08/d, so d = 1171.08 / 9325.01 = 0.12558
'That means that, for a $1 change in any stock price, the DOW will change by 1/d = 1/0.12558
'= 7.96. GE is the only stock from the "original" DOW. Suppose you had bought all 30 stocks
'in the DOW. In fact, suppose you had bought 7.96 shares of each. You 'd have paid 7.96
'(P1 + P2 + ... + P30) = 7.96(1171.08) = 9325.01. Suppose we consider a modest, annual risk-free
'rate of growth for the stock prices ... say r. (For a 2.5% rate, we'd set r = 0.025.)
'That implies a daily increase of r/360. If you held these stocks for T days, then ...
'Huh? 360? There are 365 days per year and about 250 market days, so why 360?
'There are 360 degrees in a circle. I 'm regurgitating what I read at the CBOT site. You can'
't argue with them. They're the boss.

'If you held these stocks for T days, then your stocks would be worth (Today's DOW)(1+r T/360).
'If, instead of the 30 stocks, you held a DOW future, it'd also be worth (Today's DOW)(1+r T/360)
'... except that you wouldn't get the dividends! The "Fair Value" for the DOW future would then be:
'(Today's DOW)(1+r T/360) - (T days worth of dividends)

'Note that this is the "Fair Value" at expiry, in T days. Note, too, that at expiry, T = 0 so the
'Fair Value is just the Index "Cash Value" ... less all the dividends.

'To get a theoretical Index "Cash Price" today (rather than at expiry), some investors just add a
'bunch of dividends to the current Index.. Yeah, so what's the sum of T days worth of dividends?

'Note that the average dividend yield for the 30 DOW stocks (might be!)
'I guess one could use the daily yield (= annual/360) multiplied by T.
'In any case, the difference between the "Fair Value" and the actual DOW is given
'at the CBOT site. CBOT also gives the current trading price of DOW futures
'They 're worth 9340.91, that's the "Fair Value". The actual trading value changes a lot,
'minute-to-minute. For example: They traded at anywhere from 9090 to 9425, then closed at 9298.
'Of course, Fair Value depends upon what you use for that risk-free rate, r.
'Maybe It 's some LIBOR rate, maybe it's some treasury rate.

'But suppose I'm willing to pay more than that "Fair Value"? Who'll stop me?
'Nobody. Indeed, somebody will be very happy to sell you a future which is significanlty
'greater than the Fair Value. Then that fella could go out and buy the 30 DOW stocks and make
'a killing. So if the futures deviate much from Fair Value, investors will switch to the 30
'stocks, right? Or they'll sell their basket of 30 stocks and switch to the futures. That
'arbitrage (taking advantage of the imbalance) keeps the DOW futures close to Fair Value.
'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
