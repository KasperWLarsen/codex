const CONFIG = window.CATERING_ADDIN_CONFIG;

const LUNCH_CHOICES = {
  canteen: {
    label: "Frokost i kantinen",
    requiresIndustry: false
  },
  smorrebrod: {
    label: "Smørrebrød",
    requiresIndustry: true
  },
  sandwich: {
    label: "Sandwich",
    requiresIndustry: true
  },
  poke: {
    label: "Poke bowl",
    requiresIndustry: true
  }
};

const state = {
  officeReady: false,
  previewMode: true,
  isAppointment: false,
  meeting: {
    subject: "",
    location: "",
    start: null
  }
};

const elements = {};

document.addEventListener("DOMContentLoaded", () => {
  cacheElements();
  bindEvents();
  initialize();
});

function cacheElements() {
  [
    "refreshContext",
    "meetingSubject",
    "meetingLocation",
    "meetingStart",
    "locationNotice",
    "orderForm",
    "place",
    "servingAt",
    "guestCount",
    "deadlineNotice",
    "lunchOptions",
    "lunchChoiceCanteen",
    "canteenChoiceDetail",
    "lunchChoiceSmorrebrod",
    "lunchChoiceSandwich",
    "lunchChoicePoke",
    "contactName",
    "contactEmail",
    "phone",
    "dietaryNotes",
    "comments",
    "submitOrder",
    "formStatus",
    "mailPreview",
    "mailLink",
    "previewRecipient",
    "previewSubject",
    "previewBody"
  ].forEach((id) => {
    elements[id] = document.getElementById(id);
  });
  elements.lunchChoices = Array.from(document.querySelectorAll('input[name="lunchChoice"]'));
  elements.industryChoices = [
    elements.lunchChoiceSmorrebrod,
    elements.lunchChoiceSandwich,
    elements.lunchChoicePoke
  ];
}

function bindEvents() {
  elements.refreshContext.addEventListener("click", hydrateFromOutlook);
  const handlePlaceChange = () => {
    syncPreviewLocationFromSelection();
    syncMenu();
    updateDeadlineNotice();
    updateLocationGate();
  };
  elements.place.addEventListener("change", handlePlaceChange);
  elements.place.addEventListener("input", handlePlaceChange);
  elements.lunchChoices.forEach((choice) => {
    choice.addEventListener("change", () => {
      clearFieldError(elements.lunchOptions);
      clearStatus();
    });
  });
  [elements.place, elements.servingAt, elements.guestCount, elements.contactName, elements.contactEmail].forEach((field) => {
    field.addEventListener("input", () => clearFieldError(field));
    field.addEventListener("change", () => clearFieldError(field));
  });
  elements.servingAt.addEventListener("change", updateDeadlineNotice);
  elements.guestCount.addEventListener("input", syncCountsFromGuests);
  elements.orderForm.addEventListener("submit", handleSubmit);
}

function initialize() {
  applyPreviewDefaults();
  renderMeetingContext();
  syncMenu();
  updateDeadlineNotice();
  updateLocationGate();

  if (window.Office && Office.onReady) {
    Office.onReady(async () => {
      state.officeReady = true;
      await hydrateFromOutlook();
    });
  }
}

async function hydrateFromOutlook() {
  clearStatus();

  if (!window.Office || !Office.context?.mailbox?.item) {
    state.officeReady = false;
    state.previewMode = true;
    applyPreviewDefaults();
    renderMeetingContext();
    syncMenu();
    updateDeadlineNotice();
    updateLocationGate();
    return;
  }

  const item = Office.context.mailbox.item;
  state.previewMode = false;
  const appointmentType = Office.MailboxEnums?.ItemType?.Appointment || "appointment";
  state.isAppointment = item.itemType === appointmentType || item.itemType === "appointment";
  state.meeting.subject = await getSubject(item);
  state.meeting.location = await getLocation(item);
  state.meeting.start = await getStart(item);

  if (state.meeting.start) {
    elements.servingAt.value = toDateTimeLocal(state.meeting.start);
  }

  if (Office.context.mailbox.userProfile?.emailAddress && !elements.contactEmail.value) {
    elements.contactEmail.value = Office.context.mailbox.userProfile.emailAddress;
  }

  if (Office.context.mailbox.userProfile?.displayName && !elements.contactName.value) {
    elements.contactName.value = Office.context.mailbox.userProfile.displayName;
  }

  renderMeetingContext();
  syncMenu();
  updateDeadlineNotice();
  updateLocationGate();
}

function applyPreviewDefaults() {
  state.previewMode = true;
  state.isAppointment = false;
  state.meeting.subject = "Preview";
  state.meeting.location = "Preview-mødelokale";
  state.meeting.start = addDays(new Date(), 3);
  elements.place.value = elements.place.value || "";
  elements.servingAt.value = toDateTimeLocal(state.meeting.start);
  updateCanteenChoiceDetail();
}

async function getSubject(item) {
  if (item.subject?.getAsync) {
    return getAsyncValue(item.subject);
  }

  return item.normalizedSubject || item.subject || "";
}

async function getLocation(item) {
  if (item.location?.getAsync) {
    return getAsyncValue(item.location);
  }

  return item.location || "";
}

async function getStart(item) {
  if (item.start?.getAsync) {
    return getAsyncValue(item.start);
  }

  return item.start ? new Date(item.start) : null;
}

function getAsyncValue(accessor) {
  return new Promise((resolve) => {
    accessor.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        resolve("");
      }
    });
  });
}

function renderMeetingContext() {
  elements.meetingSubject.textContent = state.meeting.subject || "Ingen titel";
  elements.meetingLocation.textContent = state.meeting.location || "Ingen lokation";
  elements.meetingStart.textContent = state.meeting.start ? formatDateTime(state.meeting.start) : "Intet tidspunkt";
}

function updateLocationGate() {
  const gate = getLocationGate();
  elements.locationNotice.className = `notice ${gate.level}`;
  elements.locationNotice.textContent = gate.message;
  elements.submitOrder.disabled = !gate.allowed;
  elements.place.disabled = false;
}

function getLocationGate() {
  if (state.previewMode) {
    return {
      allowed: true,
      level: "warn",
      message: "Previewtilstand: bestillingssted vælges manuelt. Mødets lokation medsendes kun som reference."
    };
  }

  if (!state.isAppointment) {
    return {
      allowed: false,
      level: "error",
      message: "Bestilling kan kun oprettes fra en Outlook-aftale."
    };
  }

  if (!state.meeting.location) {
    return {
      allowed: true,
      level: "warn",
      message: "Mødets lokation er tom. Vælg bestillingssted manuelt nedenfor."
    };
  }

  return {
    allowed: true,
    level: "ok",
    message: "Bestillingssted vælges manuelt. Mødets lokation medsendes kun som reference."
  };
}

function syncMenu() {
  const isIndustry = canOrderIndustryProducts();
  const hasPlace = Boolean(elements.place.value);

  updateCanteenChoiceDetail();
  elements.lunchChoiceCanteen.disabled = !hasPlace;

  elements.industryChoices.forEach((choice) => {
    choice.disabled = !isIndustry;
    choice.closest(".choice-option").hidden = !isIndustry;
  });

  const selectedChoice = getSelectedLunchChoice();
  if (!isIndustry && LUNCH_CHOICES[selectedChoice]?.requiresIndustry) {
    elements.lunchChoiceCanteen.checked = true;
  }
}

function canOrderIndustryProducts() {
  return elements.place.value === "industrivej8";
}

function getSelectedLunchChoice() {
  return elements.lunchChoices.find((choice) => choice.checked)?.value || "";
}

function getCanteenChoiceLabel() {
  const selectedPlace = CONFIG.allowedPlaces[elements.place.value];
  return selectedPlace?.canteenLabel || selectedPlace?.label || "Vælg bestillingssted først";
}

function updateCanteenChoiceDetail() {
  elements.canteenChoiceDetail.textContent = getCanteenChoiceLabel();
}

function syncPreviewLocationFromSelection() {
  if (!state.previewMode) {
    return;
  }

  updateCanteenChoiceDetail();
  renderMeetingContext();
}

function syncCountsFromGuests() {
  clearStatus();
}

function updateDeadlineNotice() {
  const servingAt = getServingDate();
  elements.deadlineNotice.className = "deadline";
  elements.deadlineNotice.textContent = "";

  if (!servingAt) {
    return;
  }

  const hoursUntilServing = (servingAt.getTime() - Date.now()) / 36e5;
  if (hoursUntilServing < 48) {
    elements.deadlineNotice.className = "deadline warn";
    elements.deadlineNotice.textContent = "Bestilling bør helst ske mindst 2 dage før levering.";
  }
}

function handleSubmit(event) {
  event.preventDefault();
  clearStatus();

  const validation = validateOrder();
  if (!validation.valid) {
    markFieldError(validation.field);
    showStatus(validation.message, "error");
    return;
  }

  const order = collectOrder();
  openOrderEmail(order);
}

function validateOrder() {
  const gate = getLocationGate();
  if (!gate.allowed) {
    return { valid: false, message: gate.message };
  }

  const requiredBeforeMenu = [
    [elements.place, "Vælg bestillingssted."],
    [elements.servingAt, "Vælg dato og tidspunkt."],
    [elements.guestCount, "Udfyld antal personer."]
  ];

  for (const [field, message] of requiredBeforeMenu) {
    if (!field.value.trim()) {
      field.focus();
      return { valid: false, message, field };
    }
  }

  if (toPositiveInt(elements.guestCount.value) < 1) {
    elements.guestCount.focus();
    return { valid: false, message: "Antal personer skal være mindst 1.", field: elements.guestCount };
  }

  const canOrderIndustry = canOrderIndustryProducts();
  const selectedChoice = getSelectedLunchChoice();

  if (!selectedChoice) {
    elements.lunchChoiceCanteen.focus();
    return { valid: false, message: "Vælg én frokostmulighed.", field: elements.lunchOptions };
  }

  if (!canOrderIndustry && LUNCH_CHOICES[selectedChoice]?.requiresIndustry) {
    elements.lunchChoiceCanteen.focus();
    return { valid: false, message: "Smørrebrød, sandwich og poke bowl kan kun vælges, når bestillingsstedet er Boardroom.", field: elements.lunchOptions };
  }

  const requiredAfterMenu = [
    [elements.contactName, "Udfyld navn."],
    [elements.contactEmail, "Udfyld e-mail."]
  ];

  for (const [field, message] of requiredAfterMenu) {
    if (!field.value.trim()) {
      field.focus();
      return { valid: false, message, field };
    }
  }

  if (!elements.contactEmail.validity.valid) {
    elements.contactEmail.focus();
    return { valid: false, message: "E-mailadressen er ikke gyldig.", field: elements.contactEmail };
  }

  return { valid: true };
}

function collectOrder() {
  const canOrderIndustry = canOrderIndustryProducts();
  const place = elements.place.value;
  const lunchChoice = getSelectedLunchChoice();
  const safeLunchChoice = canOrderIndustry || !LUNCH_CHOICES[lunchChoice]?.requiresIndustry ? lunchChoice : "canteen";
  const servingAt = getServingDate();
  const products = [[LUNCH_CHOICES[safeLunchChoice].label, toPositiveInt(elements.guestCount.value)]];

  return {
    place,
    placeLabel: CONFIG.allowedPlaces[place].label,
    lunchChoice: safeLunchChoice,
    lunchChoiceLabel: LUNCH_CHOICES[safeLunchChoice].label,
    servingAt,
    guestCount: toPositiveInt(elements.guestCount.value),
    products,
    contactName: elements.contactName.value.trim(),
    contactEmail: elements.contactEmail.value.trim(),
    phone: elements.phone.value.trim(),
    dietaryNotes: elements.dietaryNotes.value.trim(),
    comments: elements.comments.value.trim(),
    meetingSubject: state.meeting.subject,
    meetingLocation: state.meeting.location,
    meetingStart: state.meeting.start
  };
}

function openOrderEmail(order) {
  const subject = `Frokostbestilling - ${order.lunchChoiceLabel} - ${formatDate(order.servingAt)}`;
  const htmlBody = buildEmailHtml(order);
  const plainBody = buildPlainText(order);
  const mailtoUrl = buildMailtoUrl(subject, plainBody);

  if (window.Office && Office.context?.mailbox?.displayNewMessageForm) {
    try {
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: [CONFIG.recipientEmail],
        subject,
        htmlBody
      });
      showStatus("Mailen er klargjort i Outlook. Tryk Send i mailvinduet for at sende bestillingen.", "success");
      return;
    } catch (error) {
      console.warn("Outlook kunne ikke åbne en ny meddelelsesformular.", error);
    }
  }

  showMailPreview(subject, plainBody, mailtoUrl);
  try {
    window.open(mailtoUrl, "_blank");
  } catch (error) {
    console.warn("Mailprogrammet kunne ikke åbnes automatisk.", error);
  }
  showStatus("Mailudkastet er klargjort nedenfor. Brug knappen 'Åbn i mailprogram', hvis Outlook ikke åbnede automatisk.", "success");
}

function buildMailtoUrl(subject, plainBody) {
  return `mailto:${encodeURIComponent(CONFIG.recipientEmail)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(plainBody)}`;
}

function showMailPreview(subject, plainBody, mailtoUrl) {
  elements.previewRecipient.value = CONFIG.recipientEmail;
  elements.previewSubject.value = subject;
  elements.previewBody.value = plainBody;
  elements.mailLink.href = mailtoUrl;
  elements.mailPreview.hidden = false;
}

function buildEmailHtml(order) {
  const orderRows = [
    ["Bestillingssted", order.placeLabel],
    ["Frokostvalg", order.lunchChoiceLabel],
    ["Dato og tidspunkt", formatDateTime(order.servingAt)],
    ["Antal personer", order.guestCount]
  ];
  const contactRows = [
    ["Navn", order.contactName],
    ["E-mail", order.contactEmail],
    ["Telefon", order.phone || "-"]
  ];
  const noteRows = [
    ["Allergier/særlige hensyn", order.dietaryNotes || "-"],
    ["Kommentarer", order.comments || "-"]
  ];
  const meetingRows = [
    ["Mødetitel", order.meetingSubject || "-"],
    ["Mødelokation", order.meetingLocation || "-"],
    ["Mødetidspunkt", order.meetingStart ? formatDateTime(order.meetingStart) : "-"]
  ];
  const deadlineText = isInsideDeadline(order.servingAt) ? `
    <tr>
      <td style="padding:12px 16px;background:#fff3cf;color:#8a5c00;border-radius:6px;font-size:14px;">
        <strong>Bemærk:</strong> Bestillingen er oprettet mindre end 2 dage før levering/servering.
      </td>
    </tr>
  ` : "";

  const section = (title, rows) => `
    <tr>
      <td style="padding:24px 0 8px;font-size:16px;font-weight:700;color:#172020;border-top:1px solid #eef1ee;">${escapeHtml(title)}</td>
    </tr>
    <tr>
      <td>
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #d8ded9;border-radius:6px;overflow:hidden;">
          ${rows.map(([label, value]) => `
            <tr>
              <td width="36%" style="padding:10px 12px;background:#f7f8f5;border-bottom:1px solid #d8ded9;font-weight:700;color:#65706c;font-size:13px;vertical-align:top;">${escapeHtml(String(label))}</td>
              <td style="padding:10px 12px;border-bottom:1px solid #d8ded9;color:#172020;font-size:13px;vertical-align:top;">${formatEmailValue(value)}</td>
            </tr>
          `).join("")}
        </table>
      </td>
    </tr>
  `;

  return `
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="font-family:Segoe UI,Arial,sans-serif;background:#f7f8f5;padding:18px;color:#172020;">
      <tr>
        <td align="center">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:680px;background:#ffffff;border:1px solid #d8ded9;border-radius:8px;padding:0 18px 18px;">
            <tr>
              <td style="padding:22px 0 4px;color:#156064;font-size:12px;font-weight:700;text-transform:uppercase;">Frokostbestilling</td>
            </tr>
            <tr>
              <td style="padding:0 0 8px;font-size:24px;font-weight:700;color:#172020;">Ny frokostbestilling</td>
            </tr>
            ${deadlineText}
            ${section("Bestilling", orderRows)}
            ${section("Kontakt", contactRows)}
            ${section("Bemærkninger", noteRows)}
            ${section("Mødereference", meetingRows)}
          </table>
        </td>
      </tr>
    </table>
  `;
}

function buildPlainText(order) {
  const lines = [
    "========================================",
    "FROKOSTBESTILLING",
    "========================================",
    "",
    "----------------------------------------",
    "BESTILLING",
    "----------------------------------------",
    `Bestillingssted: ${order.placeLabel}`,
    `Frokostvalg: ${order.lunchChoiceLabel}`,
    `Dato og tidspunkt: ${formatDateTime(order.servingAt)}`,
    `Antal personer: ${order.guestCount}`,
    "",
    "----------------------------------------",
    "KONTAKT",
    "----------------------------------------",
    `Navn: ${order.contactName}`,
    `E-mail: ${order.contactEmail}`,
    `Telefon: ${order.phone || "-"}`,
    "",
    "----------------------------------------",
    "BEMÆRKNINGER",
    "----------------------------------------",
    `Allergier/særlige hensyn: ${order.dietaryNotes || "-"}`,
    `Kommentarer: ${order.comments || "-"}`,
    "",
    "----------------------------------------",
    "MØDEREFERENCE",
    "----------------------------------------",
    `Mødetitel: ${order.meetingSubject || "-"}`,
    `Mødelokation: ${order.meetingLocation || "-"}`,
    `Mødetidspunkt: ${order.meetingStart ? formatDateTime(order.meetingStart) : "-"}`
  ];

  if (isInsideDeadline(order.servingAt)) {
    lines.push(
      "",
      "BEMÆRK",
      "Bestillingen ligger inden for 2 dage."
    );
  }

  return lines.join("\r\n");
}

function getServingDate() {
  return elements.servingAt.value ? new Date(elements.servingAt.value) : null;
}

function isInsideDeadline(date) {
  return date && (date.getTime() - Date.now()) / 36e5 < 48;
}

function toPositiveInt(value) {
  const number = Number.parseInt(value, 10);
  return Number.isFinite(number) && number > 0 ? number : 0;
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function toDateTimeLocal(date) {
  const local = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 16);
}

function formatDate(date) {
  return new Intl.DateTimeFormat("da-DK", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric"
  }).format(date);
}

function formatDateTime(date) {
  return new Intl.DateTimeFormat("da-DK", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  }).format(date);
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatEmailValue(value) {
  return escapeHtml(String(value)).replaceAll("\n", "<br>");
}

function showStatus(message, level) {
  elements.formStatus.className = `form-status ${level}`;
  elements.formStatus.textContent = message;
}

function clearStatus() {
  elements.formStatus.className = "form-status";
  elements.formStatus.textContent = "";
  hideMailPreview();
}

function markFieldError(field) {
  clearFieldErrors();
  if (!field) {
    return;
  }
  field.classList.add("field-error");
}

function clearFieldError(field) {
  field.classList.remove("field-error");
}

function clearFieldErrors() {
  [elements.place, elements.servingAt, elements.guestCount, elements.lunchOptions, elements.contactName, elements.contactEmail].forEach(clearFieldError);
}

function hideMailPreview() {
  if (elements.mailPreview) {
    elements.mailPreview.hidden = true;
  }
}
