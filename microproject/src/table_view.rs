use egui_table::Column;

pub const SHEET_ROW_HEADER_WIDTH: f32 = 44.0;
pub const SHEET_NAME_WIDTH: f32 = 220.0;
pub const SHEET_DURATION_WIDTH: f32 = 80.0;
pub const SHEET_DATE_WIDTH: f32 = 100.0;
pub const SHEET_PERCENT_WIDTH: f32 = 80.0;
pub const SHEET_PREDECESSOR_WIDTH: f32 = 130.0;

pub const SHEET_COLUMN_WIDTHS: [f32; 7] = [
    SHEET_ROW_HEADER_WIDTH,
    SHEET_NAME_WIDTH,
    SHEET_DURATION_WIDTH,
    SHEET_DATE_WIDTH,
    SHEET_DATE_WIDTH,
    SHEET_PERCENT_WIDTH,
    SHEET_PREDECESSOR_WIDTH,
];

pub fn sheet_fixed_columns_width() -> f32 {
    SHEET_COLUMN_WIDTHS.iter().copied().sum()
}

pub fn build_egui_table_columns(chart_width: f32) -> Vec<Column> {
    let mut columns: Vec<Column> = SHEET_COLUMN_WIDTHS
        .into_iter()
        .map(|width| Column::new(width).resizable(false))
        .collect();

    columns.push(Column::new(chart_width).resizable(true));
    columns
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builds_expected_column_count() {
        let columns = build_egui_table_columns(480.0);
        assert_eq!(columns.len(), 8);
        assert_eq!(sheet_fixed_columns_width(), 754.0);
    }
}
