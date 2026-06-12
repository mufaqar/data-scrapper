import { NextRequest } from "next/server";
import * as XLSX from "xlsx";

const SODA_API = "https://data.transportation.gov/resource/kjg3-diqy.csv";

const COLUMNS = [
  "dot_number",
  "legal_name",
  "dba_name",
  "carrier_operation",
  "hm_flag",
  "pc_flag",
  "phy_street",
  "phy_city",
  "phy_state",
  "phy_zip",
  "phy_country",
  "mailing_street",
  "mailing_city",
  "mailing_state",
  "mailing_zip",
  "mailing_country",
  "telephone",
  "fax",
  "email_address",
  "mcs150_date",
  "nbr_power_unit",
  "driver_total",
  "authorized_for_hire",
  "private_only",
  "exempt_for_hire",
  "op_other",
] as const;

const HEADERS = [
  "USDOT #",
  "Legal Name",
  "DBA Name",
  "Operation",
  "HM Flag",
  "PC Flag",
  "Street",
  "City",
  "State",
  "ZIP",
  "Country",
  "Mailing Street",
  "Mailing City",
  "Mailing State",
  "Mailing ZIP",
  "Mailing Country",
  "Phone",
  "Fax",
  "Email",
  "MCS-150 Date",
  "Power Units",
  "Drivers",
  "Authorized for Hire",
  "Private Only",
  "Exempt for Hire",
  "Other Ops",
];

interface Carrier {
  dot_number: string;
  legal_name: string;
  dba_name: string;
  carrier_operation: string;
  hm_flag: string;
  pc_flag: string;
  phy_street: string;
  phy_city: string;
  phy_state: string;
  phy_zip: string;
  phy_country: string;
  mailing_street: string;
  mailing_city: string;
  mailing_state: string;
  mailing_zip: string;
  mailing_country: string;
  telephone: string;
  fax: string;
  email_address: string;
  mcs150_date: string;
  nbr_power_unit: string;
  driver_total: string;
  authorized_for_hire: string;
  private_only: string;
  exempt_for_hire: string;
  op_other: string;
}

function buildSoql(city: string, state?: string, limit = 50_000) {
  const cityUpper = city.toUpperCase().trim();
  const conditions = [
    `carrier_operation = 'A'`,
    `phy_city = '${cityUpper.replace(/'/g, "''")}'`,
  ];
  if (state) {
    conditions.push(`phy_state = '${state.toUpperCase().trim()}'`);
  }
  return `SELECT ${COLUMNS.join(",")} WHERE ${conditions.join(" AND ")} ORDER BY dot_number LIMIT ${limit}`;
}

async function fetchCarriers(city: string, state?: string): Promise<Carrier[]> {
  const soql = buildSoql(city, state);
  const url = `${SODA_API}?$query=${encodeURIComponent(soql)}`;
  const resp = await fetch(url, { headers: { Accept: "text/csv" } });
  if (!resp.ok) {
    const body = await resp.text();
    throw new Error(`SODA API error ${resp.status}: ${body.slice(0, 300)}`);
  }
  const text = await resp.text();
  const lines = text.trim().split("\n");
  if (lines.length < 2) return [];
  const csvHeaders = lines[0].replace(/^"|"$/g, "").split('","');
  const carriers: Carrier[] = [];
  for (let i = 1; i < lines.length; i++) {
    const values = lines[i].replace(/^"|"$/g, "").split('","');
    const record: Record<string, string> = {};
    csvHeaders.forEach((h, idx) => {
      record[h.trim()] = (values[idx] || "").trim();
    });
    carriers.push(record as unknown as Carrier);
  }
  return carriers;
}

function isOwnerOperator(c: Carrier, maxPowerUnits = 2): boolean {
  const pu = parseInt(c.nbr_power_unit, 10);
  if (isNaN(pu) || pu < 1 || pu > maxPowerUnits) return false;
  if (c.authorized_for_hire !== "true") return false;
  return true;
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { city, state } = body as { city?: string; state?: string };

    if (!city || !city.trim()) {
      return Response.json({ error: "City is required" }, { status: 400 });
    }

    const raw = await fetchCarriers(city, state);
    const owners = raw.filter((c) => isOwnerOperator(c));

    const wb = XLSX.utils.book_new();
    const rows = owners.map((c) => COLUMNS.map((key) => c[key] ?? ""));
    const data = [HEADERS, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(data);

    ws["!cols"] = HEADERS.map((_, i) => {
      const maxLen = data.reduce((acc, row) => {
        const val = row[i]?.toString() ?? "";
        return Math.max(acc, Math.min(val.length, 50));
      }, HEADERS[i].length);
      return { wch: maxLen + 3 };
    });

    XLSX.utils.book_append_sheet(wb, ws, "Carriers");

    const summaryRows = [
      ["Field", "Value"],
      ["Extraction Date", new Date().toISOString().split("T")[0]],
      [
        "Data Source",
        "data.transportation.gov - SMS Input - Motor Carrier Census Information",
      ],
      ["City", city.toUpperCase().trim()],
      ...(state ? [["State", state.toUpperCase().trim()]] : []),
      ["Total Carriers Found", raw.length.toString()],
      ["Owner-Operators", owners.length.toString()],
      ["Filter", `Power units 1-2, Interstate (A), Authorized for Hire`],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
    wsSummary["!cols"] = [{ wch: 25 }, { wch: 60 }];
    XLSX.utils.book_append_sheet(wb, wsSummary, "Summary");

    const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    const filename = `truck_owners_${city.trim().replace(/\s+/g, "_").toUpperCase()}${state ? `_${state.toUpperCase()}` : ""}.xlsx`;

    return new Response(buffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filename}"`,
      },
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : "Unknown error";
    console.error("Scrape error:", msg);
    return Response.json({ error: msg }, { status: 500 });
  }
}
