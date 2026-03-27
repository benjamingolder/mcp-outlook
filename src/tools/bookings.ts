import { Client } from "@microsoft/microsoft-graph-client";

export async function listBookingBusinesses(client: Client) {
  const result = await client.api("/solutions/bookingBusinesses").get();
  return result.value.map((b: any) => ({
    id: b.id,
    displayName: b.displayName,
    email: b.email,
    phone: b.phone,
    webSiteUrl: b.webSiteUrl,
    isPublished: b.isPublished,
  }));
}

export async function listBookingServices(client: Client, params: { businessId: string }) {
  const { businessId } = params;
  const result = await client.api(`/solutions/bookingBusinesses/${businessId}/services`).get();
  return result.value.map((s: any) => ({
    id: s.id,
    displayName: s.displayName,
    description: s.description,
    duration: s.defaultDuration,
    price: s.defaultPrice,
    priceType: s.defaultPriceType,
    isHiddenFromCustomers: s.isHiddenFromCustomers,
  }));
}

export async function listBookingAppointments(client: Client, params: {
  businessId: string;
  start?: string;
  end?: string;
}) {
  const { businessId, start, end } = params;

  let req = client.api(`/solutions/bookingBusinesses/${businessId}/appointments`);
  if (start && end) {
    req = req.query({ start, end });
  }

  const result = await req.get();
  return result.value.map((a: any) => ({
    id: a.id,
    serviceId: a.serviceId,
    serviceName: a.serviceName,
    start: a.startDateTime,
    end: a.endDateTime,
    customerName: a.customers?.[0]?.name ?? null,
    customerEmail: a.customers?.[0]?.emailAddress ?? null,
    duration: a.duration,
    price: a.price,
    staffMemberIds: a.staffMemberIds,
  }));
}

export async function createBookingAppointment(client: Client, params: {
  businessId: string;
  serviceId: string;
  startDateTime: string;
  endDateTime: string;
  timeZone?: string;
  customerName: string;
  customerEmail: string;
  customerPhone?: string;
  staffMemberIds?: string[];
  notes?: string;
}) {
  const {
    businessId,
    serviceId,
    startDateTime,
    endDateTime,
    timeZone = "Europe/Berlin",
    customerName,
    customerEmail,
    customerPhone,
    staffMemberIds = [],
    notes,
  } = params;

  const result = await client
    .api(`/solutions/bookingBusinesses/${businessId}/appointments`)
    .post({
      serviceId,
      startDateTime: { dateTime: startDateTime, timeZone },
      endDateTime: { dateTime: endDateTime, timeZone },
      staffMemberIds,
      ...(notes && { additionalInformation: notes }),
      customers: [
        {
          name: customerName,
          emailAddress: customerEmail,
          ...(customerPhone && { phone: customerPhone }),
        },
      ],
    });

  return { id: result.id, serviceName: result.serviceName, start: result.startDateTime };
}

export async function cancelBookingAppointment(client: Client, params: {
  businessId: string;
  appointmentId: string;
  reason?: string;
}) {
  const { businessId, appointmentId, reason = "" } = params;
  await client
    .api(`/solutions/bookingBusinesses/${businessId}/appointments/${appointmentId}/cancel`)
    .post({ cancellationMessage: reason });
  return { success: true, message: "Termin storniert." };
}
