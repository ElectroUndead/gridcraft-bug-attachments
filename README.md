# GridCraft bug report attachments

Public attachments for a bug report about sorting from a column's AutoFilter menu:
the header row is sorted into the data (GridCraft 0.6.1).

* `GridCraft-header-sort-sample.xlsx` - minimal English sample workbook (no customer data)
* `sample-before.png` / `sample-after-header-in-data.png` / `sample-workaround.png` - screenshots
* `make_sample.py`, `probe_trigger.py` - scripts that build the sample and probe the sorted range

Reproduce:

```
gridcraft-cli run --in GridCraft-header-sort-sample.xlsx --cmd 'data.sortAscending={"column":"B"}'
#   {"command":"data.sortAscending","result":{"sorted":"A2:C8"}}   <- header row inside the sorted range
```