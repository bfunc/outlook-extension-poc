// The "Key Elements" record the pane edits: one idea with a status, a headline, a trading area and
// instrument legs. The lists (trading areas, instruments) come from be-mock/samples.json.

import samples from "../be-mock/samples.json";

export type IdeaStatus = "New" | "Pipeline" | "RevEnq";
export type Side = "Buy" | "Pay" | "Receive" | "Sell";

export type Leg = {
  id: string;
  instrument: string;
  emea: boolean;
  factual: boolean;
  side: Side;
  price: string;
  underlying: string;
  timeHorizon: string;
};

export type KeyElements = {
  status: IdeaStatus;
  headline: string;
  tradingArea: string;
  legs: Leg[];
};

export const IDEA_STATUSES: IdeaStatus[] = ["New", "Pipeline", "RevEnq"];
export const SIDES: Side[] = ["Buy", "Pay", "Receive", "Sell"];
export const TRADING_AREAS: string[] = samples.tradingAreas;
export const INSTRUMENTS: { name: string; underlyingLabel: string }[] = samples.instruments;

export const DEFAULT_TIME_HORIZON = "Good when sent";

export function emptyKeyElements(): KeyElements {
  return { status: "New", headline: "", tradingArea: "", legs: [] };
}

export function newLeg(instrument: string): Leg {
  return {
    id: Math.random().toString(36).slice(2, 10),
    instrument,
    emea: false,
    factual: false,
    side: "Buy",
    price: "",
    underlying: "",
    timeHorizon: DEFAULT_TIME_HORIZON,
  };
}

/** A leg is an investment recommendation when it is EMEA-traded and not a factual market comment. */
export function isRecommendation(leg: Leg): boolean {
  return leg.emea && !leg.factual;
}

export function cardTitle(ke: KeyElements): string {
  return ke.legs.some(isRecommendation) ? "Investment recommendation" : "Idea";
}

export function underlyingLabel(instrument: string): string {
  return INSTRUMENTS.find((i) => i.name === instrument)?.underlyingLabel ?? "Underlying(s)";
}

export function parseKeyElements(json: string | undefined | null): KeyElements | null {
  if (!json) return null;
  try {
    const v = JSON.parse(json);
    return v && typeof v === "object" && Array.isArray(v.legs) ? (v as KeyElements) : null;
  } catch {
    return null;
  }
}
