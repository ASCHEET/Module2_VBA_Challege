# **Module2_VBA_Challege**

## Steves' parents are looking to invest in green energy as and alternative to fossil fuels. Steve would like to help by doing more research into green companies for his parents, he wants to expand the dataset to include the entire stock market over the last few years. The stock analysis script analyses of 2017 and 2018 stocks of the following companies: Atlantica Sustainable "AY", Canadian Solar Inc "CSIQ", DAQO New Energy Corp "DQ", Enphase Energy Inc "ENPH", First Solar Inc "FSLR", Hannon Armstrong Sustainable Infrastructure "HASI", Jinkosolar Holding Co. Ltd. "JKS", Sunrun Inc "RUN", Solaredge Technologies Inc "SEDG", Sunpower Corp "SPWR", TerraForm Power, Inc "TERP", and Vivint Solar "VSLR".

# Analysis Model

## Visual Basic for Excel was used to study the data from the green energy stocks to evaluate their performance over the course of the calendar year.  The return calculation is a comparison between the Stock Ending Price over the stock Starting Price for the given calendar year.

### Looking at Figure 1, 2017 was a good year for green companies in general; DAQO New Energy Cory "DQ" had a 199% return on stock price, Solaredge Technologies "SEDG" had a return of 185%, surely the best performers for 2017.
### ![Figure 1 - Yearly performance of Green Stocks for 2017](https://github.com/ASCHEET/Module2_VBA_Challege/blob/main/Resources/2017_analysis.png?raw=true)
### Analysis of the same green stocks for 2018 shows that the majority of green stocks listed lost returns based on the stock price for the calendar year.  However; it is important to highlight that in Figure 2 (below) the stock with the most loss is DAQO New Energy "DQ" with -63% losses and Solaredge Technologies "SEDG" only had a -8% loss.  An 8% loss in 2018, isn't that bad after such a lucrative gains in 2017 when looking over a two year period.
### ![Figure 2 - Yearly performance of Green Stocks for 2018](https://github.com/ASCHEET/Module2_VBA_Challege/blob/main/Resources/2018_analysis.png?raw=true)
### The sleeper investment appears to be Enphase Energy Inc "ENPH" that showed positive results in 2017 (130%) and 2018 (82%) as well as being one of the most highly traded companies in 2018.  Sunrun Inc "RUN" had a modest year for 2017 at 6% gains, but a breakout year in 2018, yielding 84% return and a combined rate of return of  in both years of +94%.

# Results
## It is necessary to look at both years of analysis to help determine what stock to consider a good investment.  
### Enphase Energy had a good year in 2017 showing a 130% return on stock value, but a great 2018 showing an additional 82% return during a year where green energy was mostly down or stagnate.  The highest performing stock being Enphase Energy Inc "ENPH" with an overall rate of return of 318% over both 2017 and 2018.  Enphase Energy Inc also had the highest trading volume for the year of 2018.  Singling out Enphsase, the highest performer over 2017 and 2018, had Steve's parents invested $1,000 into Enphase Energy at the beginning of 2017 it would be worth $4,176 at the end of 2018. If Steves' parents had a diversified portfolio and invested $1,000 in all of the green stock companies at the beginning of 2017, they would have an initial investment of $12,000 dollars and a present value of $17,814.  Another stock that is second to the performance of Enphase Energy is Solaredge Technologies.  They had a rate of return at +185% for 2017 and only a 8% loss in 2018 which yields a combined rate of return of 162% over the two year period.

### Refactoring analysis was used in VBA to streamline and use the software to find the bounds and analyze information in between those bounds.  The speed used to calculate each analysis is used to determine if the refactored analysis is faster, see Figures 3-6 below.
![Figure 3 - Time of original 2017 calculation](https://github.com/ASCHEET/Module2_VBA_Challege/blob/main/Resources/2017_timer_original.png?raw=true) ![Figure 4 - Time for re-factored 2017 calculation](https://github.com/ASCHEET/Module2_VBA_Challege/blob/main/Resources/2017_timer_refactored.png?raw=true)
![Figure 5 - Time of original 2018 calculation](https://github.com/ASCHEET/Module2_VBA_Challege/blob/main/Resources/2018_timer_original.png?raw=true) ![Figure 6 - Time for re-factored 2018 calculation](https://github.com/ASCHEET/Module2_VBA_Challege/blob/main/Resources/2018_timer_refactored.png?raw=true)
### Refactoring shows a slight reduction in calculation speed.  Refactoring would most likely be more evident if more data needed to be studied.



