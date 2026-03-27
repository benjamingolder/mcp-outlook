import { Client } from "@microsoft/microsoft-graph-client";

export async function listWorksheets(client: Client, params: { fileId: string; driveId?: string }) {
  const { fileId, driveId } = params;
  const base = driveId
    ? `/drives/${driveId}/items/${fileId}`
    : `/me/drive/items/${fileId}`;
  const result = await client.api(`${base}/workbook/worksheets`).get();
  return result.value.map((w: any) => ({
    id: w.id,
    name: w.name,
    position: w.position,
    visibility: w.visibility,
  }));
}

export async function getRange(client: Client, params: {
  fileId: string;
  worksheetId: string;
  address: string;
  driveId?: string;
}) {
  const { fileId, worksheetId, address, driveId } = params;
  const base = driveId
    ? `/drives/${driveId}/items/${fileId}`
    : `/me/drive/items/${fileId}`;
  const result = await client
    .api(`${base}/workbook/worksheets/${worksheetId}/range(address='${address}')`)
    .get();
  return {
    address: result.address,
    values: result.values,
    formulas: result.formulas,
    numberFormat: result.numberFormat,
    rowCount: result.rowCount,
    columnCount: result.columnCount,
  };
}

export async function updateRange(client: Client, params: {
  fileId: string;
  worksheetId: string;
  address: string;
  values: unknown[][];
  driveId?: string;
}) {
  const { fileId, worksheetId, address, values, driveId } = params;
  const base = driveId
    ? `/drives/${driveId}/items/${fileId}`
    : `/me/drive/items/${fileId}`;
  await client
    .api(`${base}/workbook/worksheets/${worksheetId}/range(address='${address}')`)
    .patch({ values });
  return { success: true, message: "Bereich aktualisiert." };
}

export async function getUsedRange(client: Client, params: {
  fileId: string;
  worksheetId: string;
  driveId?: string;
}) {
  const { fileId, worksheetId, driveId } = params;
  const base = driveId
    ? `/drives/${driveId}/items/${fileId}`
    : `/me/drive/items/${fileId}`;
  const result = await client
    .api(`${base}/workbook/worksheets/${worksheetId}/usedRange`)
    .get();
  return {
    address: result.address,
    values: result.values,
    rowCount: result.rowCount,
    columnCount: result.columnCount,
  };
}
