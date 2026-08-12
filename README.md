# AR-to-Bank Reconciliation Automation

A personal finance automation project exploring how repetitive AR-to-bank reconciliation activities can be standardized and automated.

The project compares AR Ledger and Bank Remittance data, identifies matched transactions and reconciliation exceptions, and produces a structured reconciliation output for investigation.

## Project Objective

The goal was to build a practical reconciliation prototype that can:

- Compare transactions from two financial data sources
- Match transactions using a reconciliation key
- Detect transactions existing in only one source
- Identify duplicate transactions
- Compare amount, currency, customer/payer, and transaction date
- Apply a ±1-day date tolerance
- Classify matched transactions and exceptions
- Provide exception reasons and recommended actions
- Validate the automated results against a controlled answer key

All data used in the project is synthetic and was created specifically for testing and demonstration.

## Implementations

The same reconciliation process was implemented using two approaches.

### 1. n8n + JavaScript

The original prototype was developed in n8n using JavaScript Code nodes.

The workflow:

1. Reads AR Ledger and Bank Remittance data
2. Standardizes both sources into a common structure
3. Groups transactions using a reconciliation key
4. Applies deterministic reconciliation rules
5. Identifies matches and exceptions
6. Validates results against a controlled answer key
7. Generates reconciliation and summary outputs

### 2. Excel Power Query + VBA

The reconciliation was subsequently recreated in Excel to explore how the same process could be implemented using tools commonly used in Finance teams.

**Power Query** is used to:

- Prepare the AR and Bank datasets
- Standardize column structures
- Apply appropriate data types
- Perform basic text cleaning

**VBA** is used to:

- Group transactions by reconciliation key
- Detect duplicate and one-sided transactions
- Compare reconciliation attributes
- Apply the ±1-day date tolerance
- Classify exceptions
- Generate investigation actions
- Validate the final results

A `Run Reconciliation` button executes the VBA reconciliation process using the standardized Power Query outputs.

## Reconciliation Logic

Transactions are evaluated based on:

- Reconciliation Key
- Amount
- Currency
- Customer / Payer
- Transaction Date

Structural exceptions include:

- AR Only
- Bank Only
- Duplicate AR
- Duplicate Bank
- Duplicate in Both Sources

For normal 1-to-1 transactions, the remaining reconciliation attributes are compared before the transaction is classified as either `MATCHED` or `EXCEPTION`.

## Validation Results

Both implementations were tested using the same controlled dataset.

| Metric | Result |
|---|---:|
| Total Reconciliation Keys | 1,096 |
| Matched | 365 |
| Exceptions | 731 |
| AR Only | 62 |
| Bank Only | 43 |
| Duplicate AR | 31 |
| Duplicate Bank | 38 |
| Duplicate Both | 0 |
| Amount Discrepancy | 53 |
| Currency Discrepancy | 59 |
| Customer Discrepancy | 67 |
| Date Discrepancy | 470 |

The Excel Power Query + VBA implementation reproduced the validated n8n results across all tracked reconciliation metrics.

## Project Screenshots

### Excel Reconciliation Control

The Excel implementation uses Power Query for data preparation and VBA for the reconciliation process.

![Excel Reconciliation Control](screenshots/Excel_Control.png)

### Reconciliation Output

The reconciliation output identifies matched records, structural exceptions, field-level discrepancies, exception reasons, and recommended investigation actions.

![Reconciliation Output](screenshots/Reconciliation_Output.png)

### Validation Results

The Excel Power Query + VBA implementation was validated against the same expected results used for the n8n implementation.

![Validation Results](screenshots/Validation_ALL_PASS.png)

## Technology

- Microsoft Excel
- Power Query
- VBA
- n8n
- JavaScript
- CSV
- GitHub

## Project Scope

This is a personal proof-of-concept using controlled synthetic data.

The objective is not to represent a production reconciliation system, but to demonstrate how repetitive reconciliation activities can be standardized, automated, validated, and structured for exception investigation.

The core reconciliation logic remains deterministic to maintain consistency, auditability, and explainability.
