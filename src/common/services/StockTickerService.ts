import strings from 'StockTickerCardAdaptiveCardExtensionStrings';
import { IStockQuote, IStockDataPoint, StockTimeRange } from '../models/models';

const LOG_SOURCE = '📈 StockTickerService';

const SAMPLE_QUOTE: IStockQuote = {
  symbol: 'MSFT',
  price: 189.30,
  change: 2.45,
  changePercent: 1.31,
  open: 187.20,
  high: 190.15,
  low: 186.80,
  previousClose: 186.85,
  lastUpdated: new Date()
};

function _buildSampleDaily(days: number): IStockDataPoint[] {
  const points: IStockDataPoint[] = [];
  const base = 185;
  const now = new Date();
  for (let i = days - 1; i >= 0; i--) {
    const d = new Date(now);
    d.setDate(d.getDate() - i);
    points.push({ date: d, price: parseFloat((base + Math.sin(i / 3) * 8 + i * 0.05).toFixed(2)) });
  }
  return points;
}

export class StockTickerService {
  private _apiKey: string = '';
  private quoteURL = 'https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol={symbol}&apikey={key}';
  private dailyURL = 'https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol={symbol}&apikey={key}';
  private weeklyURL = 'https://www.alphavantage.co/query?function=TIME_SERIES_WEEKLY_ADJUSTED&symbol={symbol}&apikey={key}';

  public Init(apiKey: string): void {
    this._apiKey = apiKey;
  }

  public async getQuote(symbol: string): Promise<IStockQuote> {
    let retVal: IStockQuote;
    try {
      const url = this.quoteURL.replace('{symbol}', encodeURIComponent(symbol)).replace('{key}', this._apiKey);
      //Get the data from the API
      const resp = await fetch(url);
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const json = await resp.json();
      
      //Get quote from data and check if there is a price - if not, throw error to trigger sample data
      const quote = json['Global Quote'];
      if (!quote || !quote['05. price']) {
        throw new Error(`no data for ${symbol}; using sample`);
      }

      //Assign the data to the return value
      retVal = {
        symbol: quote['01. symbol'],
        price: parseFloat(quote['05. price']),
        change: parseFloat(quote['09. change']),
        changePercent: parseFloat(quote['10. change percent'].replace('%', '')),
        open: parseFloat(quote['02. open']),
        high: parseFloat(quote['03. high']),
        low: parseFloat(quote['04. low']),
        previousClose: parseFloat(quote['08. previous close']),
        lastUpdated: new Date()
      };
    } catch (err) {
      console.error(`${LOG_SOURCE} (getQuote) - ${err}`);
      retVal = { ...SAMPLE_QUOTE, symbol: symbol || 'MSFT' };
    }
    return retVal;
  }

  public async getHistoricalData(symbol: string, range: StockTimeRange): Promise<IStockDataPoint[]> {
    let retVal: IStockDataPoint[];
    try {
      if (range === strings.Range1YShort) {
        retVal = await this._fetchWeekly(symbol);
      } else {
        retVal = await this._fetchDaily(symbol, range === strings.Range30DShort ? 30 : 7);
      }
    } catch (err) {
      console.error(`${LOG_SOURCE} (getHistoricalData) - ${err}`);
      retVal = _buildSampleDaily(range === strings.Range1YShort ? 52 : range === strings.Range30DShort ? 30 : 7);
    }
    return retVal;
  }

  private async _fetchDaily(symbol: string, days: number): Promise<IStockDataPoint[]> {
    let retVal: IStockDataPoint[];
    try {
      const url = this.dailyURL.replace('{symbol}', encodeURIComponent(symbol)).replace('{key}', this._apiKey);
      const resp = await fetch(url);
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const json = await resp.json();
      const series = json['Time Series (Daily)'];
      if (!series) throw new Error(`no data for ${symbol}; using sample`);

      //We need to map this to our IStockDataPoint format and 
      //filter it down to the number of days we want
      retVal = Object.keys(series)
        .sort((a, b) => a.localeCompare(b))
        .slice(-days)
        .map((dateStr: string) => ({
          date: new Date(dateStr + 'T00:00:00'),
          price: parseFloat((series[dateStr] as any)['4. close'])
        }));
    } catch (err) {
      console.error(`${LOG_SOURCE} (_fetchDaily) - ${err}`);
      retVal = _buildSampleDaily(days);
    }
    return retVal;
  }

  private async _fetchWeekly(symbol: string): Promise<IStockDataPoint[]> {
    let retVal: IStockDataPoint[];
    try {
      const url = this.weeklyURL.replace('{symbol}', encodeURIComponent(symbol)).replace('{key}', this._apiKey);
      const resp = await fetch(url);
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const json = await resp.json();
      const series = json['Weekly Adjusted Time Series'];
      if (!series) throw new Error(`no data for ${symbol}; using sample`);
      
      //We need to map this to our IStockDataPoint format and 
      //filter it down to the number of weeks we want
      retVal = Object.keys(series)
        .sort((a, b) => a.localeCompare(b))
        .slice(-52)
        .map((dateStr: string) => ({
          date: new Date(dateStr + 'T00:00:00'),
          price: parseFloat((series[dateStr] as any)['5. adjusted close'])
        }));
    } catch (err) {
      console.error(`${LOG_SOURCE} (_fetchWeekly) - ${err}`);
      retVal = _buildSampleDaily(52);
    }
    return retVal;
  }
}
