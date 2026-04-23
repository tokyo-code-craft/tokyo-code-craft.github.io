---
---

# Excel 重複削除

```python
import pandas as pd
df = pd.read_excel("input.xlsx")
df = df.drop_duplicates()
df.to_excel("output.xlsx", index=False)
```
