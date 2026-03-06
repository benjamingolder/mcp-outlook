import { getGraphClient } from "../graph.js";

export async function listContacts(params: { top?: number; filter?: string; search?: string }) {
  const { top = 20, filter, search } = params;
  const client = getGraphClient();

  let req = client
    .api("/me/contacts")
    .top(top)
    .select("id,displayName,emailAddresses,mobilePhone,businessPhones,jobTitle,companyName,department");

  if (filter) req = req.filter(filter);
  if (search) req = req.search(`"${search}"`);

  const result = await req.get();
  return result.value.map((c: any) => ({
    id: c.id,
    displayName: c.displayName,
    emailAddresses: c.emailAddresses,
    mobilePhone: c.mobilePhone,
    businessPhones: c.businessPhones,
    jobTitle: c.jobTitle,
    companyName: c.companyName,
    department: c.department,
  }));
}

export async function getContact(id: string) {
  const client = getGraphClient();
  const c = await client.api(`/me/contacts/${id}`).get();
  return {
    id: c.id,
    displayName: c.displayName,
    givenName: c.givenName,
    surname: c.surname,
    emailAddresses: c.emailAddresses,
    mobilePhone: c.mobilePhone,
    businessPhones: c.businessPhones,
    homePhones: c.homePhones,
    jobTitle: c.jobTitle,
    companyName: c.companyName,
    department: c.department,
    officeLocation: c.officeLocation,
    businessAddress: c.businessAddress,
    birthday: c.birthday,
    personalNotes: c.personalNotes,
  };
}

export async function createContact(params: {
  givenName: string;
  surname?: string;
  emailAddresses?: { address: string; name?: string }[];
  mobilePhone?: string;
  businessPhones?: string[];
  jobTitle?: string;
  companyName?: string;
  department?: string;
}) {
  const client = getGraphClient();
  const result = await client.api("/me/contacts").post(params);
  return { id: result.id, displayName: result.displayName };
}

export async function updateContact(params: {
  id: string;
  givenName?: string;
  surname?: string;
  emailAddresses?: { address: string; name?: string }[];
  mobilePhone?: string;
  businessPhones?: string[];
  jobTitle?: string;
  companyName?: string;
  department?: string;
  personalNotes?: string;
}) {
  const { id, ...patch } = params;
  const client = getGraphClient();
  await client.api(`/me/contacts/${id}`).patch(patch);
  return { success: true, message: "Kontakt aktualisiert." };
}

export async function deleteContact(id: string) {
  const client = getGraphClient();
  await client.api(`/me/contacts/${id}`).delete();
  return { success: true, message: "Kontakt gelöscht." };
}
