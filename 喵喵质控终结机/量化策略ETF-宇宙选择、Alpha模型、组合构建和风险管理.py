#region imports
from AlgorithmImports import *
#endregion

def ResetAndWarmUp(algorithm, security, resolution, lookback = None):
    indicator = security['logr']
    consolidator = security['consolidator']

    if not lookback:
        lookback = indicator.WarmUpPeriod

    # historical request to update the consolidator that will warm up the indicator
    history = algorithm.History[consolidator.InputType](security.Symbol, lookback, resolution,
        dataNormalizationMode = DataNormalizationMode.ScaledRaw)

    indicator.Reset()
    
    # Replace the consolidator, since we cannot reset it
    # Not ideal since we don't the consolidator type and period
    algorithm.SubscriptionManager.RemoveConsolidator(security.Symbol, consolidator)
    consolidator = TradeBarConsolidator(timedelta(1))
    algorithm.RegisterIndicator(security.Symbol, indicator, consolidator)
    
    for bar in list(history)[:-1]:
        consolidator.Update(bar)

    return consolidator

'''
# In main.py, OnData and call HandleCorporateActions for framework models (if necessary)
    def OnData(self, slice):
        if slice.Splits or slice.Dividends:
            self.alpha.HandleCorporateActions(self, slice)
            self.pcm.HandleCorporateActions(self, slice)
            self.risk.HandleCorporateActions(self, slice)
            self.execution.HandleCorporateActions(self, slice)

# In the framework models, add
from utils import ResetAndWarmUp

and implement HandleCorporateActions. E.g.:
    def HandleCorporateActions(self, algorithm, slice):
        for security.Symbol, data in self.security.Symbol_data.items():
            if slice.Splits.ContainsKey(security.Symbol) or slice.Dividends.ContainsKey(security.Symbol):
                data.WarmUpIndicator()

where WarmUpIndicator will call ResetAndWarmUp for each indicator/consolidator pair
'''