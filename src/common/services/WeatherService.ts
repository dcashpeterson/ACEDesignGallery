import { SPFx, spfi, SPFI } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import "@pnp/sp/items";
import { IWeatherLocation, IWeatherData, WeatherLocationFields } from '../models/models';
import { ConfigService } from "./ConfigService";
import { Lists } from "../models/enums";

export class WeatherService {
  private _sp!: SPFI;
  private _configService = new ConfigService();
  private LOG_SOURCE = '🌤️ WeatherService';

  public async Init(siteUrl: string, pageContext: any): Promise<void> {
    try {
      this._sp = spfi(siteUrl).using(SPFx({ pageContext }));
      await this._configService.Init(siteUrl, pageContext);
      const created = await this._configService._configList(Lists.WEATHERLOCATIONS, 'Weather locations used by the Weather Card', WeatherLocationFields);
      if (created) {
        await this.seedSampleLocations();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (Init) - ${err}`);
    }
  }

  public async getLocations(): Promise<IWeatherLocation[]> {
    const retVal: IWeatherLocation[] = [];
    try {
      const items = await this._sp.web.lists.getByTitle(Lists.WEATHERLOCATIONS).items
        .select('Id', 'Location_x0020_Name', 'Latitude', 'Longitude')
        .top(500)();
      items.forEach(item => {
        if (item.Location_x0020_Name) {
          retVal.push({
            id: String(item.Id),
            name: item.Location_x0020_Name,
            latitude: item.Latitude ?? 0,
            longitude: item.Longitude ?? 0
          });
        }
      });
      retVal.sort((a, b) => a.name.localeCompare(b.name));
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getLocations) - ${err}`);
    }
    return retVal;
  }

  public async getWeather(latitude: number, longitude: number, unit: string): Promise<IWeatherData | undefined> {
    try {
      const tempUnit = unit === 'celsius' ? 'celsius' : 'fahrenheit';
      const url = `https://api.open-meteo.com/v1/forecast?latitude=${latitude}&longitude=${longitude}&current_weather=true&temperature_unit=${tempUnit}`;
      const response = await fetch(url);
      if (!response.ok) {
        throw new Error(`Open-Meteo API error: ${response.status}`);
      }
      const json = await response.json();
      const cw = json.current_weather;
      const isDay = cw.is_day === 1;
      const code: number = cw.weathercode;
      return {
        temperature: cw.temperature,
        weatherCode: code,
        windspeed: cw.windspeed,
        isDay,
        description: this._mapWeatherCode(code),
        iconUrl: this._getWeatherIconUrl(code, isDay)
      };
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getWeather) - ${err}`);
      return undefined;
    }
  }

  public async getWeatherForAllLocations(locations: IWeatherLocation[], unit: string): Promise<{ [name: string]: IWeatherData }> {
    const result: { [name: string]: IWeatherData } = {};
    try {
      await Promise.all(locations.map(async loc => {
        const weather = await this.getWeather(loc.latitude, loc.longitude, unit);
        if (weather) {
          result[loc.name] = weather;
        }
      }));
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getWeatherForAllLocations) - ${err}`);
    }
    return result;
  }

  private _mapWeatherCode(code: number): string {
    if (code === 0) return 'Clear sky';
    if (code === 1) return 'Mainly clear';
    if (code === 2) return 'Partly cloudy';
    if (code === 3) return 'Overcast';
    if (code === 45 || code === 48) return 'Foggy';
    if (code >= 51 && code <= 57) return 'Drizzle';
    if (code >= 61 && code <= 67) return 'Rain';
    if (code >= 71 && code <= 77) return 'Snow';
    if (code >= 80 && code <= 82) return 'Rain showers';
    if (code === 85 || code === 86) return 'Snow showers';
    if (code === 95) return 'Thunderstorm';
    if (code === 96 || code === 99) return 'Thunderstorm with hail';
    return 'Unknown';
  }

  private _getWeatherIconUrl(code: number, isDay: boolean): string {
    const suffix = isDay ? 'd' : 'n';
    let icon: string;
    if (code === 0) {
      icon = `01${suffix}`;
    } else if (code === 1) {
      icon = `02${suffix}`;
    } else if (code === 2) {
      icon = `03${suffix}`;
    } else if (code === 3) {
      icon = `04${suffix}`;
    } else if (code === 45 || code === 48) {
      icon = `50${suffix}`;
    } else if (code >= 51 && code <= 57) {
      icon = `09${suffix}`;
    } else if (code >= 61 && code <= 67) {
      icon = `10${suffix}`;
    } else if (code >= 71 && code <= 77) {
      icon = `13${suffix}`;
    } else if (code >= 80 && code <= 82) {
      icon = `09${suffix}`;
    } else if (code === 85 || code === 86) {
      icon = `13${suffix}`;
    } else if (code >= 95) {
      icon = `11${suffix}`;
    } else {
      icon = `01${suffix}`;
    }
    return `https://openweathermap.org/img/wn/${icon}@2x.png`;
  }

  public async seedSampleLocations(): Promise<void> {
    try {
      const list = this._sp.web.lists.getByTitle(Lists.WEATHERLOCATIONS);
      const samples = [
        { Title: 'New York',     Location_x0020_Name: 'New York',     Latitude: 40.7128,  Longitude: -74.0060 },
        { Title: 'London',       Location_x0020_Name: 'London',       Latitude: 51.5074,  Longitude: -0.1278  },
        { Title: 'Tokyo',        Location_x0020_Name: 'Tokyo',        Latitude: 35.6762,  Longitude: 139.6503 },
        { Title: 'Sydney',       Location_x0020_Name: 'Sydney',       Latitude: -33.8688, Longitude: 151.2093 },
        { Title: 'Los Angeles',  Location_x0020_Name: 'Los Angeles',  Latitude: 34.0522,  Longitude: -118.2437 }
      ];
      for (const loc of samples) {
        await list.items.add(loc);
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (_seedSampleLocations) - ${err}`);
    }
  }
}
