// app.js - Vue SPA logic
const { createApp } = Vue;

function formatCurrency(value) {
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(value);
}

const app = createApp({
  data() {
    return {
      taxRate: 0.07,
      products: [
        { id: 1, name: 'Solar Panel 300W', category: 'Solar', price: 200, qty: 0 },
        { id: 2, name: 'Inverter 5kW', category: 'Inverter', price: 1500, qty: 0 },
        { id: 3, name: 'Battery 10kWh', category: 'Battery', price: 8000, qty: 0 },
      ],
    };
  },
  computed: {
    subtotal() {
      return this.products.reduce((sum, p) => sum + p.price * p.qty, 0);
    },
    tax() {
      return this.subtotal * this.taxRate;
    },
    grandTotal() {
      return this.subtotal + this.tax;
    },
  },
  methods: {
    formatCurrency,
    saveDraft() {
      localStorage.setItem('quotationDraft', JSON.stringify(this.products));
      alert('Draft saved locally');
    },
    clearDraft() {
      this.products.forEach(p => (p.qty = 0));
      localStorage.removeItem('quotationDraft');
    },
    loadDraft() {
      const saved = localStorage.getItem('quotationDraft');
      if (saved) {
        const arr = JSON.parse(saved);
        this.products.forEach(p => {
          const match = arr.find(item => item.id === p.id);
          if (match) p.qty = match.qty;
        });
      }
    },
    downloadPdf(type) {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF();
      doc.setFontSize(16);
      doc.text('Aye Yeik Nyo Co., Ltd.', 10, 20);
      doc.text(type.charAt(0).toUpperCase() + type.slice(1) + ' Document', 10, 30);
      let y = 40;
      this.products.forEach(p => {
        if (p.qty > 0) {
          doc.text(`${p.name} (${p.qty} x ${formatCurrency(p.price)}) = ${formatCurrency(p.qty * p.price)}`, 10, y);
          y += 10;
        }
      });
      doc.text(`Subtotal: ${formatCurrency(this.subtotal)}`, 10, y);
      y += 10;
      doc.text(`Tax (${this.taxRate * 100}%): ${formatCurrency(this.tax)}`, 10, y);
      y += 10;
      doc.text(`Grand Total: ${formatCurrency(this.grandTotal)}`, 10, y);
      doc.save(`${type}-${Date.now()}.pdf`);
    },
  },
  mounted() {
    this.loadDraft();
  },
});

app.mount('#app');
