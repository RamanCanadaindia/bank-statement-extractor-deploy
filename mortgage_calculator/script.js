const form = document.querySelector("#mortgage-form");
const errorBox = document.querySelector("#error-box");

const fields = {
  income: document.querySelector("#income"),
  gdsr: document.querySelector("#gdsr"),
  propertyTax: document.querySelector("#property-tax"),
  heatingCost: document.querySelector("#heating-cost"),
  condoFee: document.querySelector("#condo-fee"),
  contractRate: document.querySelector("#contract-rate"),
  amortization: document.querySelector("#amortization"),
  downPayment: document.querySelector("#down-payment")
};

const outputs = {
  qualifyingRate: document.querySelector("#qualifying-rate"),
  monthlyPayment: document.querySelector("#monthly-payment"),
  maxLoan: document.querySelector("#max-loan"),
  purchasePrice: document.querySelector("#purchase-price")
};

const moneyFormatter = new Intl.NumberFormat("en-CA", {
  style: "currency",
  currency: "CAD",
  maximumFractionDigits: 0
});

function readNumber(input, fallback = 0) {
  if (input.value.trim() === "") {
    return fallback;
  }

  return Number(input.value);
}

function setError(message) {
  errorBox.textContent = message;
  errorBox.hidden = false;
}

function clearError() {
  errorBox.textContent = "";
  errorBox.hidden = true;
}

function resetResults() {
  outputs.qualifyingRate.textContent = "-";
  outputs.monthlyPayment.textContent = "-";
  outputs.maxLoan.textContent = "-";
  outputs.purchasePrice.textContent = "-";
}

function calculateMortgage(event) {
  event.preventDefault();
  clearError();

  const values = {
    income: readNumber(fields.income, NaN),
    gdsr: readNumber(fields.gdsr, 39),
    propertyTax: readNumber(fields.propertyTax),
    heatingCost: readNumber(fields.heatingCost),
    condoFee: readNumber(fields.condoFee),
    contractRate: readNumber(fields.contractRate),
    amortization: readNumber(fields.amortization, 25),
    downPayment: readNumber(fields.downPayment)
  };

  if (fields.income.value.trim() === "" || Number.isNaN(values.income)) {
    resetResults();
    setError("Please enter gross annual income.");
    fields.income.focus();
    return;
  }

  const hasNegativeValue = Object.values(values).some((value) => value < 0);
  if (hasNegativeValue) {
    resetResults();
    setError("Please enter zero or positive numbers only.");
    return;
  }

  if (values.amortization <= 0) {
    resetResults();
    setError("Amortization years must be greater than 0.");
    fields.amortization.focus();
    return;
  }

  const qualifyingRate = Math.max(values.contractRate + 2, 5.25);
  const annualCondoFeePortion = values.condoFee * 12 * 0.5;
  const annualGdsrAllowance = values.income * (values.gdsr / 100);
  const monthlyPaymentAllowed = (
    annualGdsrAllowance -
    values.propertyTax -
    values.heatingCost -
    annualCondoFeePortion
  ) / 12;

  if (monthlyPaymentAllowed <= 0) {
    resetResults();
    setError("The maximum monthly mortgage payment is less than or equal to 0. Try a higher income, lower property costs, or lower GDSR inputs.");
    return;
  }

  const monthlyRate = qualifyingRate / 100 / 12;
  const numberOfPayments = values.amortization * 12;
  const maxLoan = monthlyPaymentAllowed * (1 - Math.pow(1 + monthlyRate, -numberOfPayments)) / monthlyRate;
  const purchasePrice = maxLoan + values.downPayment;

  outputs.qualifyingRate.textContent = `${qualifyingRate.toFixed(2)}%`;
  outputs.monthlyPayment.textContent = moneyFormatter.format(monthlyPaymentAllowed);
  outputs.maxLoan.textContent = moneyFormatter.format(maxLoan);
  outputs.purchasePrice.textContent = moneyFormatter.format(purchasePrice);
}

form.addEventListener("submit", calculateMortgage);

form.addEventListener("reset", () => {
  window.setTimeout(() => {
    clearError();
    resetResults();
  }, 0);
});
