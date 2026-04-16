use calamine::{Reader, open_workbook_auto};
use rust_xlsxwriter::Workbook;
use serde::Serialize;
use std::fs;
use std::path::PathBuf;
use std::time::Instant;

#[derive(Serialize)]
struct ResultRow {
    label: String,
    real: f64,
    real_min: f64,
    real_stddev: f64,
    rss_delta_kb: u64,
    iterations: usize,
    #[serde(skip_serializing_if = "Option::is_none")]
    size: Option<u64>,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let rows = int_env("RBXL_BENCH_ROWS", 5000);
    let cols = int_env("RBXL_BENCH_COLS", 10);
    let warmup = int_env("RBXL_BENCH_WARMUP", 1);
    let iterations = int_env("RBXL_BENCH_ITERATIONS", 5);
    let read_path = std::env::var("RBXL_BENCH_READ_PATH").ok();

    let (header, body) = build_dataset(rows, cols);
    let tmp_dir = std::env::temp_dir().join(format!("rbxl-rust-bench-{}", std::process::id()));
    fs::create_dir_all(&tmp_dir)?;

    let write_path = tmp_dir.join("rust_xlsxwriter.xlsx");
    let mut results = Vec::new();

    let mut write_result = measure("rust_xlsxwriter write", warmup, iterations, || {
        write_with_rust_xlsxwriter(&write_path, &header, &body)
    })?;
    write_result.size = Some(fs::metadata(&write_path)?.len());
    results.push(write_result);

    let target_read_path = read_path
        .map(PathBuf::from)
        .unwrap_or_else(|| write_path.clone());
    results.push(measure("calamine read", warmup, iterations, || {
        read_with_calamine(&target_read_path).map(|_| ())
    })?);

    println!("{}", serde_json::to_string(&results)?);
    let _ = fs::remove_dir_all(&tmp_dir);
    Ok(())
}

fn int_env(name: &str, fallback: usize) -> usize {
    std::env::var(name)
        .ok()
        .and_then(|value| value.parse::<usize>().ok())
        .unwrap_or(fallback)
}

fn build_dataset(rows: usize, cols: usize) -> (Vec<String>, Vec<Vec<CellValue>>) {
    let header = (0..cols)
        .map(|i| format!("col_{}", i + 1))
        .collect::<Vec<_>>();
    let body = (0..rows)
        .map(|row| {
            (0..cols)
                .map(|col| match col % 4 {
                    0 => CellValue::Number(row as f64),
                    1 => CellValue::Text(format!("row-{row}-col-{col}")),
                    2 => CellValue::Bool((row + col) % 2 == 1),
                    _ => CellValue::Number(((row * 100) + col) as f64 / 10.0),
                })
                .collect::<Vec<_>>()
        })
        .collect::<Vec<_>>();
    (header, body)
}

fn measure<F>(
    label: &str,
    warmup: usize,
    iterations: usize,
    mut func: F,
) -> Result<ResultRow, Box<dyn std::error::Error>>
where
    F: FnMut() -> Result<(), Box<dyn std::error::Error>>,
{
    for _ in 0..warmup {
        func()?;
    }

    let mut samples = Vec::with_capacity(iterations);
    let mut rss_deltas = Vec::with_capacity(iterations);
    for _ in 0..iterations {
        let before = rss_kb();
        let started = Instant::now();
        func()?;
        samples.push(started.elapsed().as_secs_f64());
        let after = rss_kb();
        rss_deltas.push(after.saturating_sub(before));
    }

    let mean = samples.iter().sum::<f64>() / samples.len() as f64;
    let real_min = samples
        .iter()
        .copied()
        .fold(f64::INFINITY, f64::min);
    let variance = samples
        .iter()
        .map(|sample| {
            let diff = sample - mean;
            diff * diff
        })
        .sum::<f64>()
        / samples.len() as f64;

    Ok(ResultRow {
        label: label.to_string(),
        real: mean,
        real_min,
        real_stddev: variance.sqrt(),
        rss_delta_kb: rss_deltas.into_iter().max().unwrap_or(0),
        iterations,
        size: None,
    })
}

fn rss_kb() -> u64 {
    let Ok(status) = fs::read_to_string("/proc/self/status") else {
        return 0;
    };
    status
        .lines()
        .find_map(|line| {
            let value = line.strip_prefix("VmRSS:")?;
            value.split_whitespace().next()?.parse::<u64>().ok()
        })
        .unwrap_or(0)
}

fn write_with_rust_xlsxwriter(
    path: &PathBuf,
    header: &[String],
    body: &[Vec<CellValue>],
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Bench")?;

    for (col, value) in header.iter().enumerate() {
        worksheet.write_string(0, col as u16, value)?;
    }

    for (row_index, row) in body.iter().enumerate() {
        for (col_index, value) in row.iter().enumerate() {
            let row_num = (row_index + 1) as u32;
            let col_num = col_index as u16;
            match value {
                CellValue::Text(text) => {
                    worksheet.write_string(row_num, col_num, text)?;
                }
                CellValue::Number(number) => {
                    worksheet.write_number(row_num, col_num, *number)?;
                }
                CellValue::Bool(flag) => {
                    worksheet.write_boolean(row_num, col_num, *flag)?;
                }
            }
        }
    }

    workbook.save(path)?;
    Ok(())
}

fn read_with_calamine(path: &PathBuf) -> Result<usize, Box<dyn std::error::Error>> {
    let mut workbook = open_workbook_auto(path)?;
    let range = workbook.worksheet_range("Bench")?;
    let mut count = 0;
    for row in range.rows() {
        count += row.len();
    }
    Ok(count)
}

enum CellValue {
    Text(String),
    Number(f64),
    Bool(bool),
}
