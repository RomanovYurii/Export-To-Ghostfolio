import { GhostfolioOrderType } from "./ghostfolioOrderType";

export class GhostfolioActivity {
    accountId: string;
    comment: string;
    fee: number;
    quantity: number;
    type: GhostfolioOrderType;
    unitPrice: number;
    currency: string;
    // Optional: Ghostfolio resolves the default data source for
    // non-investment types (FEE, INTEREST, LIABILITY) when it is omitted.
    dataSource?: string;
    date: string;
    symbol: string;
    tags: string[];
}
