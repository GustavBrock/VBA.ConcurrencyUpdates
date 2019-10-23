# VBA.ConcurrencyUpdates

### Automatic handling of concurrent updates using DAO in VBA

Version 1.0.2

If several users try to update the same record simultaneously, an error pops up asking what to do. That's fine, users know what to do. Contrary, if two processes driven from code do the same, there is no one to handle the situation, and it fails. That's bad.  

Here is a method to avoid this situation. 

![General](https://raw.githubusercontent.com/GustavBrock/VBA.ConcurrencyUpdates/master/images/EEconcurrency 2019.png)


### Running a test

A typical output from the test function **ConcurrencyAwareTest** is here:

```
First process
-----------------------
 7             54449.09 
    Update     54449.38     Microsoft Access has stopped the process ...
    Edit       54449.48     Microsoft Access has stopped the process ...
    Edit       54449.54     Microsoft Access has stopped the process ...
               54450.02      929           2 

Second process
-----------------------
 1             54448.38 
    Update     54448.57     Microsoft Access has stopped the process ...
    Edit       54448.69     Microsoft Access has stopped the process ...
    Edit       54449.07     Microsoft Access has stopped the process ...
               54449.29      914           2 
 2             54449.3 
               54449.54      238           1 
 3             54449.56 
    Update     54450.05     Microsoft Access has stopped the process ...
    Edit       54450.18     Microsoft Access has stopped the process ...
               54450.39      828           2 
 4             54450.4 
               54450.64      234           1 
```

Full documentation is found here:

![EE Logo](https://raw.githubusercontent.com/GustavBrock/VBA.ConcurrencyUpdates/master/images/EE%20Logo.png)

[Handle concurrent update conflicts in Access silently](https://www.experts-exchange.com/articles/25780/Handle-concurrent-update-conflicts-in-Access-silently.html)

<hr>

*If you wish to support my work or need extended support or advice, feel free to:*

<p>

[<img src="https://raw.githubusercontent.com/GustavBrock/VBA.ConcurrencyUpdates/master/images/BuyMeACoffee.png">](https://www.buymeacoffee.com/gustav/)