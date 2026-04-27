import { getSP, getSPForWeb } from '../core/sp';

export interface IWebSummary {
  Title: string;
  ServerRelativeUrl: string;
}

export class WebsService {
  async getCurrentWeb(): Promise<IWebSummary> {
    const sp = getSP();
    const w = await sp.web.select('Title', 'ServerRelativeUrl')();
    return {
      Title: String((w as { Title?: string }).Title ?? ''),
      ServerRelativeUrl: String((w as { ServerRelativeUrl?: string }).ServerRelativeUrl ?? ''),
    };
  }

  async getDirectSubsites(webServerRelativeUrl?: string): Promise<IWebSummary[]> {
    const sp = getSPForWeb(webServerRelativeUrl ?? undefined);
    const rows = (await sp.web.webs.select('Title', 'ServerRelativeUrl')()) as IWebSummary[];
    return Array.isArray(rows) ? rows : [];
  }
}
