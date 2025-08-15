let suppliers = JSON.parse(localStorage.getItem('suppliers') || '[]');
let selectedSupplierIndex = null;

const supplierNameInput = document.getElementById('supplierName');
const addSupplierBtn = document.getElementById('addSupplierBtn');
const supplierList = document.getElementById('supplierList');
const invoiceSection = document.getElementById('invoice-section');
const selectedSupplierTitle = document.getElementById('selectedSupplier');
const invoiceDateInput = document.getElementById('invoiceDate');
const invoiceAmountInput = document.getElementById('invoiceAmount');
const invoiceNoteInput = document.getElementById('invoiceNote');
const addInvoiceBtn = document.getElementById('addInvoiceBtn');
const invoiceTableBody = document.querySelector('#invoiceTable tbody');
const totalAmountSpan = document.getElementById('totalAmount');

function renderSuppliers() {
  supplierList.innerHTML = '';
  suppliers.forEach((s, i) => {
    const li = document.createElement('li');
    li.textContent = s.name;
    li.onclick = () => selectSupplier(i);
    supplierList.appendChild(li);
  });
}

function selectSupplier(index) {
  selectedSupplierIndex = index;
  invoiceSection.style.display = 'block';
  selectedSupplierTitle.textContent = suppliers[index].name;
  renderInvoices();
}

function renderInvoices() {
  const supplier = suppliers[selectedSupplierIndex];
  invoiceTableBody.innerHTML = '';
  let total = 0;
  supplier.invoices.forEach(inv => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${inv.date}</td><td>${inv.amount}</td><td>${inv.note}</td>`;
    invoiceTableBody.appendChild(tr);
    total += parseFloat(inv.amount);
  });
  totalAmountSpan.textContent = total.toFixed(2);
}

addSupplierBtn.onclick = () => {
  const name = supplierNameInput.value.trim();
  if (!name) return;
  suppliers.push({ name, invoices: [] });
  localStorage.setItem('suppliers', JSON.stringify(suppliers));
  supplierNameInput.value = '';
  renderSuppliers();
};

addInvoiceBtn.onclick = () => {
  if (selectedSupplierIndex === null) return;
  const date = invoiceDateInput.value;
  const amount = parseFloat(invoiceAmountInput.value);
  const note = invoiceNoteInput.value.trim();
  if (!date || isNaN(amount)) return;
  suppliers[selectedSupplierIndex].invoices.push({ date, amount, note });
  localStorage.setItem('suppliers', JSON.stringify(suppliers));
  invoiceDateInput.value = '';
  invoiceAmountInput.value = '';
  invoiceNoteInput.value = '';
  renderInvoices();
};

renderSuppliers();