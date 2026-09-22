# Excel Formulas Used

## 1. COUNTIF

Counts products belonging to a specific category.

```excel
=COUNTIF(C:C,A2)
=COUNTIFS(C:C,A2,D:D,B2)
=MAX(B2:B4)
=INDEX(A2:A4,MATCH(MAX(B2:B4),B2:B4,0))
