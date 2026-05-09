use crate::mspdi::ChartRange;

pub fn chart_content_width(range: ChartRange, zoom_px_per_day: f32) -> f32 {
    (range.days() as f32 * zoom_px_per_day).max(360.0)
}

pub fn chart_width(range: ChartRange, zoom_px_per_day: f32, available_width: f32, fixed_columns_width: f32) -> f32 {
    let content_width = chart_content_width(range, zoom_px_per_day);
    content_width.max((available_width - fixed_columns_width).max(420.0))
}

pub fn table_width(column_widths: &[f32]) -> f32 {
    column_widths.iter().copied().sum()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::mspdi::ChartRange;
    use chrono::NaiveDate;

    #[test]
    fn chart_width_uses_view_and_content() {
        let range = ChartRange {
            start: NaiveDate::from_ymd_opt(2025, 1, 1).unwrap(),
            end: NaiveDate::from_ymd_opt(2025, 1, 10).unwrap(),
        };
        assert!(chart_content_width(range, 20.0) >= 360.0);
        assert!(chart_width(range, 20.0, 800.0, 500.0) >= 420.0);
        assert_eq!(table_width(&[1.0, 2.0, 3.0]), 6.0);
    }
}
