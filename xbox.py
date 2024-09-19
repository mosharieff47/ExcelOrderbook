import xlwings as xw
import time
import websocket
import threading
import json
from sklearn.svm import SVC
import numpy as np

class CBPro(threading.Thread):

    bids = {}
    asks = {}

    def __init__(self):
        threading.Thread.__init__(self)
        self.url = 'wss://ws-feed.exchange.coinbase.com'
        self.tickers = ['BTC-USD','ETH-USD','LTC-USD','ADA-USD','LINK-USD','ZEC-USD','DOGE-USD','SOL-USD']
        self.depth = 5
        self.tbids = {tick:int(time.time()) for tick in self.tickers}
        self.tasks = {tick:int(time.time()) for tick in self.tickers}
        self.cbids = {tick:0 for tick in self.tickers}
        self.casks = {tick:0 for tick in self.tickers}


    def run(self):
        msg = {'type':'subscribe','product_ids': self.tickers, 'channels':['level2_batch']}
        conn = websocket.create_connection(self.url)
        conn.send(json.dumps(msg))
        oldbids, oldasks = {}, {}
        T0 = int(time.time())
        while True:
            resp = json.loads(conn.recv())
            self.parsebook(resp)
            if 'type' in resp.keys():
                if resp['type'] == 'l2update':
                    ticker = resp['product_id']
                    self.count_volume(ticker)

            if int(time.time()) - T0 > 60:
                conn.send(json.dumps({'event':'ping'}))
                T0 = int(time.time())

    def prepare_for_svm(self, ticker):
        depth = 100
        bids = self.bids[ticker]
        asks = self.asks[ticker]
        bids = list(sorted(bids.items(), reverse=True))[:depth]
        asks = list(sorted(asks.items()))[:depth]
        prices = [b[0] for b in bids][::-1] + [a[0] for a in asks]
        return prices
            
    def prepare_for_excel(self, excel, ticker, depth, head=4):
        
        bids = self.bids[ticker]
        asks = self.asks[ticker]
        bids = list(sorted(bids.items(), reverse=True))[:depth]
        asks = list(sorted(asks.items()))[:depth]
        self.depth = depth

        '''
        for i in range(depth):
            excel.range(f'B{head + i}').value = bids[i][0]
            excel.range(f'C{head + i}').value = bids[i][1]

            excel.range(f'E{head + i}').value = asks[i][0]
            excel.range(f'F{head + i}').value = asks[i][1]
        '''

        excel.range(f'B{head}:C{head+depth}').value = bids
        excel.range(f'E{head}:F{head+depth}').value = asks

        midmarket = (bids[0][0] + asks[0][0])/2.0
        return midmarket
    
    def count_volume(self, ticker):

        if int(time.time()) - self.tbids[ticker] > 1:
            self.tbids[ticker] = int(time.time())
            self.cbids[ticker] = 0
            self.tasks[ticker] = int(time.time())
            self.casks[ticker] = 0
        else:
            self.cbids[ticker] += 1
            self.casks[ticker] += 1

    def parsebook(self, resp):
        if 'type' in resp.keys():
            if resp['type'] == 'snapshot':
                ticker = resp['product_id']
                self.bids[ticker] = {float(i):float(j) for (i, j) in resp['bids']}
                self.asks[ticker] = {float(i):float(j) for (i, j) in resp['asks']}
            if resp['type'] == 'l2update':
                ticker = resp['product_id']
                for (side, price, volume) in resp['changes']:
                    price, volume = float(price), float(volume)
                    if side == 'buy':
                        if volume == 0:
                            if price in self.bids[ticker].keys():
                                del self.bids[ticker][price]
                        else:
                            self.bids[ticker][price] = volume
                            
                    if side == 'sell':
                        if volume == 0:
                            if price in self.asks[ticker].keys():
                                del self.asks[ticker][price]
                        else:
                            self.asks[ticker][price] = volume
                

class ML:

    sync = False

    def __init__(self, length=8):
        self.dataset = []
        self.output = []
        self.length = length
        self.oldprice = None
        
    def __call__(self, inputs, price):
        self.dataset.append(inputs)
        if self.oldprice != None:
            if price/self.oldprice - 1 > 0:
                self.output.append(1.0)
            else:
                self.output.append(0.0)
        self.oldprice = price
        if len(self.dataset) > self.length:
            del self.dataset[0]
            del self.output[0]
            self.sync = True

    def machine_learning(self):
        X = np.array(self.dataset)
        m, n = X.shape
        mu = (1/m)*np.ones(m).dot(X)
        cv = (1/(m-1))*(X - mu).T.dot(X - mu)
        sd = np.sqrt(np.diag(cv))
        Z = ((X - mu)/sd).tolist()
        Z = [[j if np.isnan(j) == False else 0 for j in i] for i in Z]
        svm = SVC(kernel='linear', probability=True)
        svm.fit(Z[:-1], self.output)
        O = np.array(Z[-1])
        pred = svm.predict(O.reshape(1, -1))
        prob = svm.predict_proba(O.reshape(1, -1))
        classification = 'Buy' if pred[0] == 0 else 'Sell'
        probability = prob[0][0] if classification == 'Buy' else prob[0][1]
        return classification, probability
        
        
        


cbpro = CBPro()
length = 10
MLX = {tick:ML(length=length) for tick in cbpro.tickers}
cbpro.start()

sheet = xw.Book('EXbox360.xlsm').sheets[0]
print('Booted.........')

ignite = True
olddepth = 1
kxp = 0
while ignite:
    if len(cbpro.bids) == 8:
        ticker = sheet.range("J4").value
        depth = int(sheet.range("J3").value)
        if depth != olddepth:
            sheet.range('B4:C1000').value = None
            sheet.range('E4:F1000').value = None
            olddepth = depth
        midmarket = cbpro.prepare_for_excel(sheet, ticker, depth)
        prices = cbpro.prepare_for_svm(ticker)
        MLX[ticker](prices, midmarket)
        if MLX[ticker].sync == True:
            try:
                cs, cp = MLX[ticker].machine_learning()
                sheet.range('J21').value = cs
                sheet.range('J22').value = cp
                kxp = 0
            except:
                sheet.range('J21').value = '-'
                sheet.range('J22').value = '-'
        else:
            kxp += 1
            sheet.range('J21').value = f'Wait {length - kxp}'
            sheet.range('J22').value = f'Wait {length - kxp}'
            
            
        sheet.range('J8').value = cbpro.cbids[ticker]
        sheet.range('J9').value = cbpro.casks[ticker]
        


cbpro.join()



