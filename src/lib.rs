use calamine::{Reader, open_workbook, Xlsx};
use pyo3::prelude::*;

#[pyfunction]
fn read_excel(file_path: &str) -> PyResult<Vec<Vec<String>>> {
    let mut workbook: Xlsx<_> = open_workbook(file_path).map_err(|e| {
        pyo3::exceptions::PyIOError::new_err(format!("Failed to open workbook: {}", e))
    })?;
    let mut result = Vec::new();

    if let Ok(sheet) = workbook.worksheet_range("Sheet1") {  // TODO: 240717 シート名を指定可能にせよ。
        for row in sheet.rows() {
            let mut row_data = Vec::new();  // TODO: 240717 skiprows のようなものを付けよ。
            for cell in row {
                row_data.push(cell.to_string());
            }
            result.push(row_data);
        }
        Ok(result)
    } else {
        Err(pyo3::exceptions::PyIOError::new_err("Sheet not found"))
    }
}

#[pymodule]
fn xlfr(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(read_excel, m)?)?;
    Ok(())
}