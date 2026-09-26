from __future__ import annotations

import sqlite3
from collections import Counter
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Iterator

from .calculator import distribution_status, guide_target
from .config import DEFAULT_TAX_RATE
from .excel_reader import normalize_text
from .models import ParsedSource, Product, ProductResult, RunCalculation, RunSummary
from .ordering import default_article_order, insert_at_group_end
from .tax_rates import TaxRatePeriod, TaxRateSchedule


DOUBLE_COUNT_EVENT = "Двойной учёт акта"
TAX_RATE_EVENT = "Налоговая ставка"


@dataclass(slots=True)
class RunReplacement:
    """Пересчитанный отчет, который заменяет сохраненный с тем же периодом."""

    old_run_id: int
    calculation: RunCalculation
    stored_paths: dict[str, Path]
    created_at: str
    planned_prices: dict[str, float] = field(default_factory=dict)


def _normalized_accrual_type(value: str) -> str:
    return normalize_text(value).replace("ё", "е")


class Database:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()
        self._seed_defaults()
        self._ensure_tax_rates()
        self._ensure_product_order()
        self._ensure_run_names()

    def connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA journal_mode = WAL")
        return connection

    @contextmanager
    def read(self) -> Iterator[sqlite3.Connection]:
        connection = self.connect()
        try:
            yield connection
        finally:
            connection.close()

    @contextmanager
    def transaction(self) -> Iterator[sqlite3.Connection]:
        connection = self.connect()
        try:
            connection.execute("BEGIN")
            yield connection
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def _initialize(self) -> None:
        with self.transaction() as db:
            db.executescript(
                """
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS products (
                    article TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    category TEXT NOT NULL DEFAULT '',
                    material_cost REAL NOT NULL DEFAULT 0,
                    labor_cost REAL NOT NULL DEFAULT 0,
                    active INTEGER NOT NULL DEFAULT 1,
                    sort_order INTEGER,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS product_cost_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    changed_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    article TEXT NOT NULL,
                    old_name TEXT,
                    new_name TEXT NOT NULL,
                    old_category TEXT,
                    new_category TEXT NOT NULL DEFAULT '',
                    old_material_cost REAL,
                    new_material_cost REAL NOT NULL,
                    old_labor_cost REAL,
                    new_labor_cost REAL NOT NULL,
                    old_active INTEGER,
                    new_active INTEGER NOT NULL,
                    change_source TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS ix_product_cost_history_article
                    ON product_cost_history(article, changed_at DESC);

                CREATE TABLE IF NOT EXISTS runs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    report_name TEXT,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    period_start TEXT,
                    period_end TEXT,
                    tax_rate REAL NOT NULL,
                    source_count INTEGER NOT NULL,
                    units REAL NOT NULL,
                    revenue REAL NOT NULL,
                    financial_result REAL NOT NULL,
                    cost_sold REAL NOT NULL,
                    tax REAL NOT NULL,
                    net_profit REAL NOT NULL,
                    unallocated_total REAL NOT NULL,
                    taxable_unallocated_income REAL NOT NULL DEFAULT 0,
                    unallocated_income_tax REAL NOT NULL DEFAULT 0,
                    tax_period_start TEXT,
                    tax_period_end TEXT,
                    realization_revenue REAL NOT NULL DEFAULT 0,
                    realization_units REAL NOT NULL DEFAULT 0,
                    duplicate_realization_rows INTEGER NOT NULL DEFAULT 0,
                    already_accrued_realization_rows INTEGER NOT NULL DEFAULT 0,
                    status TEXT NOT NULL DEFAULT 'Готов'
                );

                CREATE TABLE IF NOT EXISTS source_files (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
                    original_name TEXT NOT NULL,
                    original_path TEXT NOT NULL,
                    stored_path TEXT NOT NULL,
                    file_hash TEXT NOT NULL,
                    report_type TEXT NOT NULL,
                    sheet_name TEXT NOT NULL,
                    header_row INTEGER NOT NULL,
                    row_count INTEGER NOT NULL,
                    total_amount REAL NOT NULL,
                    period_start TEXT,
                    period_end TEXT
                );
                CREATE INDEX IF NOT EXISTS ix_source_files_hash ON source_files(file_hash);
                CREATE INDEX IF NOT EXISTS ix_source_files_run ON source_files(run_id);

                CREATE TABLE IF NOT EXISTS product_results (
                    run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
                    article TEXT NOT NULL,
                    name TEXT NOT NULL,
                    category TEXT NOT NULL DEFAULT '',
                    material_cost REAL NOT NULL,
                    labor_cost REAL NOT NULL,
                    units REAL NOT NULL,
                    revenue_no_points REAL NOT NULL,
                    partner_programs REAL NOT NULL,
                    points REAL NOT NULL,
                    commission REAL NOT NULL,
                    processing REAL NOT NULL,
                    delivery REAL NOT NULL,
                    logistics REAL NOT NULL,
                    reverse_logistics REAL NOT NULL,
                    returns_cancels REAL NOT NULL,
                    acquiring REAL NOT NULL,
                    stars REAL NOT NULL,
                    packaging REAL NOT NULL,
                    compensation REAL NOT NULL,
                    other REAL NOT NULL,
                    financial_result REAL NOT NULL,
                    tax REAL NOT NULL DEFAULT 0,
                    PRIMARY KEY (run_id, article)
                );

                CREATE TABLE IF NOT EXISTS unallocated (
                    run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
                    accrual_type TEXT NOT NULL,
                    row_count INTEGER NOT NULL,
                    amount REAL NOT NULL,
                    PRIMARY KEY (run_id, accrual_type)
                );

                CREATE TABLE IF NOT EXISTS accrual_stats (
                    run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
                    normalized_type TEXT NOT NULL,
                    accrual_type TEXT NOT NULL,
                    with_article INTEGER NOT NULL,
                    without_article INTEGER NOT NULL,
                    PRIMARY KEY (run_id, normalized_type)
                );

                CREATE TABLE IF NOT EXISTS quality_events (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
                    severity TEXT NOT NULL,
                    event_type TEXT NOT NULL,
                    message TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS scenario_prices (
                    run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
                    article TEXT NOT NULL,
                    planned_price REAL NOT NULL,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY (run_id, article)
                );

                -- Справочник ставок: valid_from IS NULL — строка «с начала».
                CREATE TABLE IF NOT EXISTS tax_rates (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    valid_from TEXT,
                    rate REAL NOT NULL CHECK (rate >= 0 AND rate <= 1)
                );
                CREATE UNIQUE INDEX IF NOT EXISTS ux_tax_rates_valid_from
                    ON tax_rates(COALESCE(valid_from, ''));

                -- Снимок справочника, по которому рассчитан отчет.
                CREATE TABLE IF NOT EXISTS run_tax_rates (
                    run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
                    position INTEGER NOT NULL,
                    valid_from TEXT,
                    rate REAL NOT NULL,
                    PRIMARY KEY (run_id, position)
                );
                """
            )
            product_columns = {row[1] for row in db.execute("PRAGMA table_info(products)")}
            if "sort_order" not in product_columns:
                db.execute("ALTER TABLE products ADD COLUMN sort_order INTEGER")
            if "category" not in product_columns:
                db.execute("ALTER TABLE products ADD COLUMN category TEXT NOT NULL DEFAULT ''")
            history_columns = {row[1] for row in db.execute("PRAGMA table_info(product_cost_history)")}
            if "old_category" not in history_columns:
                db.execute("ALTER TABLE product_cost_history ADD COLUMN old_category TEXT")
            if "new_category" not in history_columns:
                db.execute(
                    "ALTER TABLE product_cost_history ADD COLUMN new_category TEXT NOT NULL DEFAULT ''"
                )
            result_columns = {row[1] for row in db.execute("PRAGMA table_info(product_results)")}
            if "category" not in result_columns:
                db.execute(
                    "ALTER TABLE product_results ADD COLUMN category TEXT NOT NULL DEFAULT ''"
                )
            run_columns = {row[1] for row in db.execute("PRAGMA table_info(runs)")}
            if "report_name" not in run_columns:
                db.execute("ALTER TABLE runs ADD COLUMN report_name TEXT")
            if "taxable_unallocated_income" not in run_columns:
                db.execute(
                    "ALTER TABLE runs ADD COLUMN "
                    "taxable_unallocated_income REAL NOT NULL DEFAULT 0"
                )
                db.execute(
                    """
                    UPDATE runs
                    SET taxable_unallocated_income = COALESCE((
                        SELECT SUM(CASE WHEN amount > 0 THEN amount ELSE 0 END)
                        FROM unallocated
                        WHERE unallocated.run_id = runs.id
                    ), 0)
                    """
                )
            # До 0.5.30 налог не хранился и считался как «база × ставка отчета».
            # Переносим ровно эти значения, чтобы результаты старых отчетов не изменились.
            if "tax" not in result_columns:
                db.execute("ALTER TABLE product_results ADD COLUMN tax REAL NOT NULL DEFAULT 0")
                db.execute(
                    """
                    UPDATE product_results
                    SET tax = (revenue_no_points + partner_programs) * (
                        SELECT tax_rate FROM runs WHERE runs.id = product_results.run_id
                    )
                    """
                )
            if "unallocated_income_tax" not in run_columns:
                db.execute(
                    "ALTER TABLE runs ADD COLUMN unallocated_income_tax REAL NOT NULL DEFAULT 0"
                )
                db.execute("UPDATE runs SET unallocated_income_tax = taxable_unallocated_income * tax_rate")
            if "tax_period_start" not in run_columns:
                db.execute("ALTER TABLE runs ADD COLUMN tax_period_start TEXT")
                db.execute("ALTER TABLE runs ADD COLUMN tax_period_end TEXT")
                self._backfill_tax_periods(db)
            # Старый отчет рассчитан по одной ставке на все даты.
            db.execute(
                """
                INSERT INTO run_tax_rates(run_id, position, valid_from, rate)
                SELECT id, 0, NULL, tax_rate FROM runs
                WHERE id NOT IN (SELECT run_id FROM run_tax_rates)
                """
            )

    @staticmethod
    def _backfill_tax_periods(db: sqlite3.Connection) -> None:
        dates_by_run: dict[int, list[str]] = {}
        for row in db.execute("SELECT id, period_start, period_end FROM runs"):
            dates_by_run[int(row[0])] = [value for value in (row[1], row[2]) if value]
        for row in db.execute("SELECT run_id, period_start, period_end FROM source_files"):
            dates_by_run.setdefault(int(row[0]), []).extend(
                value for value in (row[1], row[2]) if value
            )
        for run_id, values in dates_by_run.items():
            if values:
                db.execute(
                    "UPDATE runs SET tax_period_start = ?, tax_period_end = ? WHERE id = ?",
                    (min(values), max(values), run_id),
                )

    def _ensure_tax_rates(self) -> None:
        """Создать справочник из прежней единой ставки настроек (одна строка «с начала»)."""
        with self.transaction() as db:
            if db.execute("SELECT COUNT(*) FROM tax_rates").fetchone()[0]:
                return
            row = db.execute("SELECT value FROM settings WHERE key = 'tax_rate'").fetchone()
            try:
                rate = float(row[0]) if row else DEFAULT_TAX_RATE
            except (TypeError, ValueError):
                rate = DEFAULT_TAX_RATE
            if not 0 <= rate <= 1:
                rate = DEFAULT_TAX_RATE
            db.execute("INSERT INTO tax_rates(valid_from, rate) VALUES (NULL, ?)", (rate,))

    def tax_rate_schedule(self) -> TaxRateSchedule:
        with self.read() as db:
            return _read_tax_schedule(db)

    def replace_tax_rates(self, schedule: TaxRateSchedule) -> None:
        with self.transaction() as db:
            self._write_tax_schedule(db, schedule)

    @staticmethod
    def _write_tax_schedule(db: sqlite3.Connection, schedule: TaxRateSchedule) -> None:
        db.execute("DELETE FROM tax_rates")
        db.executemany(
            "INSERT INTO tax_rates(valid_from, rate) VALUES (?, ?)",
            [(_date_text(item.valid_from), item.rate) for item in schedule.periods],
        )

    def run_tax_schedule(self, run_id: int) -> TaxRateSchedule:
        with self.read() as db:
            return _read_run_tax_schedule(db, run_id)

    def _seed_defaults(self) -> None:
        defaults = {
            "theme": "system",
            "tax_rate": str(DEFAULT_TAX_RATE),
            "warn_without_realization": "1",
            "preview_rows": "500",
        }
        with self.transaction() as db:
            for key, value in defaults.items():
                db.execute("INSERT OR IGNORE INTO settings(key, value) VALUES (?, ?)", (key, value))

    def _ensure_product_order(self) -> None:
        with self.transaction() as db:
            rows = db.execute("SELECT article, sort_order FROM products").fetchall()
            if not rows or all(row["sort_order"] is not None for row in rows):
                return
            positioned = sorted(
                (row for row in rows if row["sort_order"] is not None),
                key=lambda row: (int(row["sort_order"]), str(row["article"]).casefold()),
            )
            order = [str(row["article"]) for row in positioned]
            missing = [str(row["article"]) for row in rows if row["sort_order"] is None]
            if not order:
                order = default_article_order(missing)
            else:
                for article in default_article_order(missing):
                    order = insert_at_group_end(order, article)
            self._write_product_order(db, order)

    def _ensure_run_names(self) -> None:
        with self.transaction() as db:
            rows = db.execute(
                "SELECT id, report_name, period_start, period_end FROM runs"
            ).fetchall()
            for row in rows:
                current_name = str(row["report_name"] or "").strip()
                legacy_name = f"Отчет Ozon #{int(row['id'])}"
                if current_name and current_name != legacy_name:
                    continue
                db.execute(
                    "UPDATE runs SET report_name = ? WHERE id = ?",
                    (_default_run_name(int(row["id"]), row["period_start"], row["period_end"]), row["id"]),
                )

    def get_setting(self, key: str, default: str = "") -> str:
        with self.read() as db:
            row = db.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
        return str(row[0]) if row else default

    def set_setting(self, key: str, value: str) -> None:
        with self.transaction() as db:
            db.execute(
                "INSERT INTO settings(key, value) VALUES (?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                (key, value),
            )

    def list_products(self, active_only: bool = False) -> list[Product]:
        sql = "SELECT article, name, category, material_cost, labor_cost, active, sort_order FROM products"
        parameters: tuple[object, ...] = ()
        if active_only:
            sql += " WHERE active = 1"
        sql += " ORDER BY sort_order, article COLLATE NOCASE"
        with self.read() as db:
            rows = db.execute(sql, parameters).fetchall()
        return [
            Product(
                article=row["article"],
                name=row["name"],
                material_cost=float(row["material_cost"]),
                labor_cost=float(row["labor_cost"]),
                active=bool(row["active"]),
                sort_order=int(row["sort_order"]) if row["sort_order"] is not None else None,
                category=str(row["category"] or ""),
            )
            for row in rows
        ]

    def product_map(self, active_only: bool = False) -> dict[str, Product]:
        return {product.article: product for product in self.list_products(active_only=active_only)}

    def save_product(self, product: Product, source: str = "Ручное изменение") -> None:
        self.save_products([product], source=source)

    def save_products(self, products: list[Product], source: str) -> int:
        if not products:
            return 0
        articles: set[str] = set()
        for product in products:
            if not product.article.strip():
                raise ValueError("Артикул не может быть пустым")
            if product.article.strip() in articles:
                raise ValueError(f"Артикул {product.article} повторяется в импорте")
            articles.add(product.article.strip())
            if product.material_cost < 0 or product.labor_cost < 0:
                raise ValueError("Себестоимость не может быть отрицательной")
        changed = 0
        new_articles: list[str] = []
        with self.transaction() as db:
            for product in products:
                article = product.article.strip()
                name = product.name.strip() or article
                old = db.execute(
                    "SELECT name, category, material_cost, labor_cost, active FROM products WHERE article = ?",
                    (article,),
                ).fetchone()
                is_changed = old is None or (
                    old["name"] != name
                    or str(old["category"] or "") != product.category.strip()
                    or abs(float(old["material_cost"]) - product.material_cost) >= 0.005
                    or abs(float(old["labor_cost"]) - product.labor_cost) >= 0.005
                    or bool(old["active"]) != product.active
                )
                if not is_changed:
                    continue
                if old is None:
                    new_articles.append(article)
                changed += 1
                db.execute(
                    """
                    INSERT INTO product_cost_history(
                        article, old_name, new_name, old_category, new_category,
                        old_material_cost, new_material_cost,
                        old_labor_cost, new_labor_cost, old_active, new_active, change_source
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        article,
                        old["name"] if old else None,
                        name,
                        str(old["category"] or "") if old else None,
                        product.category.strip(),
                        float(old["material_cost"]) if old else None,
                        product.material_cost,
                        float(old["labor_cost"]) if old else None,
                        product.labor_cost,
                        int(old["active"]) if old else None,
                        1 if product.active else 0,
                        source,
                    ),
                )
                db.execute(
                    """
                    INSERT INTO products(article, name, category, material_cost, labor_cost, active)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT(article) DO UPDATE SET
                        name = excluded.name,
                        category = excluded.category,
                        material_cost = excluded.material_cost,
                        labor_cost = excluded.labor_cost,
                        active = excluded.active,
                        updated_at = CURRENT_TIMESTAMP
                    """,
                    (
                        article,
                        name,
                        product.category.strip(),
                        product.material_cost,
                        product.labor_cost,
                        1 if product.active else 0,
                    ),
                )
                # Категория — текущая классификация товара, а не расчетный показатель.
                # Обновляем ее во всей истории, не затрагивая сохраненную себестоимость и суммы.
                db.execute(
                    "UPDATE product_results SET category = ? WHERE article = ?",
                    (product.category.strip(), article),
                )
            if new_articles:
                current = [
                    str(row[0])
                    for row in db.execute(
                        "SELECT article FROM products WHERE sort_order IS NOT NULL ORDER BY sort_order, article COLLATE NOCASE"
                    )
                ]
                for article in default_article_order(new_articles):
                    current = insert_at_group_end(current, article)
                self._write_product_order(db, current)
        return changed

    def reorder_products(self, articles: list[str]) -> None:
        if len(articles) != len(set(articles)):
            raise ValueError("В порядке товаров повторяется артикул")
        with self.transaction() as db:
            existing = {str(row[0]) for row in db.execute("SELECT article FROM products")}
            if set(articles) != existing:
                raise ValueError("Порядок должен содержать все товары справочника")
            self._write_product_order(db, articles)

    def clear_products(self) -> int:
        """Delete the editable product catalog without changing saved report snapshots."""
        with self.transaction() as db:
            count = int(db.execute("SELECT COUNT(*) FROM products").fetchone()[0])
            db.execute("DELETE FROM products")
        return count

    @staticmethod
    def _write_product_order(db: sqlite3.Connection, articles: list[str]) -> None:
        db.executemany(
            "UPDATE products SET sort_order = ? WHERE article = ?",
            [(index, article) for index, article in enumerate(articles, start=1)],
        )

    def list_product_cost_history(self, limit: int = 500) -> list[dict[str, object]]:
        with self.read() as db:
            rows = db.execute(
                """
                SELECT changed_at, article, old_name, new_name,
                       old_category, new_category,
                       old_material_cost, new_material_cost,
                       old_labor_cost, new_labor_cost,
                       old_active, new_active, change_source
                FROM product_cost_history
                ORDER BY id DESC LIMIT ?
                """,
                (limit,),
            ).fetchall()
        return [dict(row) for row in rows]

    def find_runs_by_hash(self, file_hash: str) -> list[int]:
        with self.read() as db:
            rows = db.execute(
                "SELECT DISTINCT run_id FROM source_files WHERE file_hash = ? ORDER BY run_id DESC",
                (file_hash,),
            ).fetchall()
        return [int(row[0]) for row in rows]

    def find_runs_by_period(self, period_start: date, period_end: date) -> list[int]:
        with self.read() as db:
            rows = db.execute(
                """
                SELECT id
                FROM runs
                WHERE period_start = ? AND period_end = ?
                ORDER BY id ASC
                """,
                (_date_text(period_start), _date_text(period_end)),
            ).fetchall()
        return [int(row[0]) for row in rows]

    def save_run(
        self,
        calculation: RunCalculation,
        stored_paths: dict[str, Path],
        replace_run_ids: list[int] | None = None,
    ) -> int:
        with self.transaction() as db:
            run_id, replaced_paths = self._insert_run(
                db, calculation, stored_paths, replace_run_ids
            )
        self._remove_unreferenced_source_files(replaced_paths)
        return run_id

    def _insert_run(
        self,
        db: sqlite3.Connection,
        calculation: RunCalculation,
        stored_paths: dict[str, Path],
        replace_run_ids: list[int] | None = None,
    ) -> tuple[int, list[str]]:
        """Записать отчет в открытую транзакцию; вернуть id и файлы замененных отчетов."""
        totals = calculation.totals()
        replace_ids = sorted(set(replace_run_ids or []))
        replaced_paths: list[str] = []
        preserved_name = ""
        if replace_ids:
            placeholders = ",".join("?" for _ in replace_ids)
            replaced = db.execute(
                f"""
                SELECT id, report_name, period_start, period_end
                FROM runs
                WHERE id IN ({placeholders})
                ORDER BY id ASC
                """,
                replace_ids,
            ).fetchall()
            if len(replaced) != len(replace_ids):
                raise KeyError("Обновляемый отчет уже удален")
            expected_period = (
                _date_text(calculation.period_start),
                _date_text(calculation.period_end),
            )
            if any(
                (row["period_start"], row["period_end"]) != expected_period
                for row in replaced
            ):
                raise ValueError(
                    "Можно заменять только отчет с тем же периодом"
                )
            preserved_name = str(replaced[0]["report_name"] or "").strip()
            replaced_paths = [
                str(row["stored_path"])
                for row in db.execute(
                    f"SELECT stored_path FROM source_files WHERE run_id IN ({placeholders})",
                    replace_ids,
                ).fetchall()
            ]
        cursor = db.execute(
            """
            INSERT INTO runs(
                period_start, period_end, tax_rate, source_count, units, revenue,
                financial_result, cost_sold, tax, net_profit, unallocated_total,
                taxable_unallocated_income, unallocated_income_tax,
                tax_period_start, tax_period_end, realization_revenue, realization_units,
                duplicate_realization_rows, already_accrued_realization_rows, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Готов')
            """,
            (
                _date_text(calculation.period_start),
                _date_text(calculation.period_end),
                calculation.tax_rate,
                len(calculation.source_files),
                totals["units"],
                totals["revenue"],
                totals["financial_result"],
                totals["cost_sold"],
                totals["tax"],
                totals["net_profit"],
                calculation.unallocated_total,
                calculation.taxable_unallocated_income,
                calculation.unallocated_income_tax,
                _date_text(calculation.tax_period_start),
                _date_text(calculation.tax_period_end),
                calculation.realization_revenue,
                calculation.realization_units,
                calculation.duplicate_realization_rows,
                calculation.already_accrued_realization_rows,
            ),
        )
        run_id = int(cursor.lastrowid)
        db.execute(
            "UPDATE runs SET report_name = ? WHERE id = ?",
            (
                preserved_name
                or _default_run_name(
                    run_id,
                    _date_text(calculation.period_start),
                    _date_text(calculation.period_end),
                ),
                run_id,
            ),
        )
        for source in calculation.source_files:
            db.execute(
                """
                INSERT INTO source_files(
                    run_id, original_name, original_path, stored_path, file_hash,
                    report_type, sheet_name, header_row, row_count, total_amount,
                    period_start, period_end
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_id,
                    source.path.name,
                    str(source.path),
                    str(stored_paths[str(source.path)]),
                    source.file_hash,
                    source.report_type,
                    source.sheet_name,
                    source.header_row,
                    source.row_count,
                    source.total_amount,
                    _date_text(source.period_start),
                    _date_text(source.period_end),
                ),
            )
        for item in calculation.products:
            db.execute(
                """
                INSERT INTO product_results(
                    run_id, article, name, category, material_cost, labor_cost, units,
                    revenue_no_points, partner_programs, points, commission, processing,
                    delivery, logistics, reverse_logistics, returns_cancels, acquiring,
                    stars, packaging, compensation, other, financial_result, tax
                ) VALUES (
                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
                )
                """,
                _product_result_tuple(run_id, item) + (item.tax(calculation.tax_rate),),
            )
        schedule = calculation.tax_schedule or TaxRateSchedule.single(calculation.tax_rate)
        db.executemany(
            "INSERT INTO run_tax_rates(run_id, position, valid_from, rate) VALUES (?, ?, ?, ?)",
            [
                (run_id, position, _date_text(item.valid_from), item.rate)
                for position, item in enumerate(schedule.periods)
            ],
        )
        for accrual_type, (row_count, amount) in calculation.unallocated.items():
            db.execute(
                "INSERT INTO unallocated VALUES (?, ?, ?, ?)",
                (run_id, accrual_type, row_count, amount),
            )
        for accrual_type, (with_article, without_article) in calculation.accrual_stats.items():
            db.execute(
                "INSERT INTO accrual_stats VALUES (?, ?, ?, ?, ?)",
                (run_id, _normalized_accrual_type(accrual_type), accrual_type, with_article, without_article),
            )
        for article, message in calculation.skipped_articles.items():
            db.execute(
                "INSERT INTO quality_events(run_id, severity, event_type, message) VALUES (?, 'Предупреждение', 'Пропущенный артикул', ?)",
                (run_id, message),
            )
        for sku in sorted(calculation.sku_conflicts):
            db.execute(
                "INSERT INTO quality_events(run_id, severity, event_type, message) VALUES (?, 'Предупреждение', 'Конфликт SKU', ?)",
                (run_id, f"SKU {sku} связан с несколькими артикулами"),
            )
        if calculation.duplicate_realization_rows:
            db.execute(
                "INSERT INTO quality_events(run_id, severity, event_type, message) VALUES (?, 'Информация', 'Дубли выкупов', ?)",
                (run_id, f"Пропущено одинаковых строк: {calculation.duplicate_realization_rows}"),
            )
        if calculation.already_accrued_realization_rows:
            db.execute(
                "INSERT INTO quality_events(run_id, severity, event_type, message) VALUES (?, 'Информация', 'Выручка уже начислена', ?)",
                (run_id, f"Не добавлено повторно строк: {calculation.already_accrued_realization_rows}"),
            )
        if not any(source.report_type == "REALIZATION" for source in calculation.source_files):
            db.execute(
                "INSERT INTO quality_events(run_id, severity, event_type, message) VALUES (?, 'Предупреждение', 'Нет отчета о выкупах', ?)",
                (run_id, "Расчет выполнен без RealizationReportCIS; выручка может быть неполной"),
            )
        for message in calculation.source_period_warnings:
            db.execute(
                "INSERT INTO quality_events(run_id, severity, event_type, message) "
                "VALUES (?, 'Предупреждение', 'Несовпадение периодов', ?)",
                (run_id, message),
            )
        for message in calculation.double_count_warnings:
            db.execute(
                "INSERT INTO quality_events(run_id, severity, event_type, message) "
                "VALUES (?, 'Предупреждение', ?, ?)",
                (run_id, DOUBLE_COUNT_EVENT, message),
            )
        for message in calculation.tax_rate_warnings:
            db.execute(
                "INSERT INTO quality_events(run_id, severity, event_type, message) "
                "VALUES (?, 'Предупреждение', ?, ?)",
                (run_id, TAX_RATE_EVENT, message),
            )
        hash_counts = Counter(source.file_hash for source in calculation.source_files)
        run_names = {
            int(row["id"]): str(row["report_name"])
            for row in db.execute("SELECT id, report_name FROM runs").fetchall()
        }
        for source in calculation.source_files:
            previous_run_ids = [
                value for value in source.duplicate_run_ids if value not in replace_ids
            ]
            if previous_run_ids or hash_counts[source.file_hash] > 1:
                details = (
                    "ранее использовался в "
                    + ", ".join(
                        f"«{run_names[value]}»" if value in run_names else "сохраненном отчете"
                        for value in previous_run_ids
                    )
                    if previous_run_ids
                    else "повторно выбран в текущем запуске"
                )
                db.execute(
                    "INSERT INTO quality_events(run_id, severity, event_type, message) VALUES (?, 'Предупреждение', 'Повторный файл', ?)",
                    (run_id, f"{source.path.name}: {details}"),
                )
        if replace_ids:
            placeholders = ",".join("?" for _ in replace_ids)
            db.execute(f"DELETE FROM runs WHERE id IN ({placeholders})", replace_ids)
        return run_id, replaced_paths

    def replace_runs(
        self,
        replacements: list[RunReplacement],
        tax_schedule: TaxRateSchedule | None = None,
    ) -> dict[int, int]:
        """Заменить отчеты пересчитанными (и, при необходимости, справочник ставок) одной транзакцией.

        Название, дата расчета и плановые цены прежнего отчета сохраняются.
        """
        old_to_new: dict[int, int] = {}
        removed_paths: list[str] = []
        with self.transaction() as db:
            if tax_schedule is not None:
                self._write_tax_schedule(db, tax_schedule)
            for item in replacements:
                new_id, replaced_paths = self._insert_run(
                    db,
                    item.calculation,
                    item.stored_paths,
                    replace_run_ids=[item.old_run_id],
                )
                db.execute(
                    "UPDATE runs SET created_at = ? WHERE id = ?",
                    (item.created_at, new_id),
                )
                db.executemany(
                    "INSERT INTO scenario_prices(run_id, article, planned_price) VALUES (?, ?, ?)",
                    [(new_id, article, price) for article, price in item.planned_prices.items()],
                )
                removed_paths.extend(replaced_paths)
                old_to_new[item.old_run_id] = new_id
        self._remove_unreferenced_source_files(removed_paths)
        return old_to_new

    def run_tax_periods(self) -> dict[int, tuple[date | None, date | None]]:
        """Интервал дат налоговой базы каждого сохраненного отчета."""
        with self.read() as db:
            rows = db.execute("SELECT id, tax_period_start, tax_period_end FROM runs").fetchall()
        return {
            int(row["id"]): (_parse_date(row["tax_period_start"]), _parse_date(row["tax_period_end"]))
            for row in rows
        }

    def list_runs(self) -> list[RunSummary]:
        with self.read() as db:
            rows = db.execute(
                """
                SELECT r.id, r.created_at, r.period_start, r.period_end, r.source_count,
                       r.units, r.revenue, r.net_profit, r.unallocated_total, r.status,
                       r.report_name, r.unallocated_income_tax,
                       CASE WHEN r.cost_sold = 0 THEN 0
                            ELSE r.net_profit / r.cost_sold
                       END AS profitability,
                       CASE WHEN r.revenue = 0 THEN 0
                            ELSE -COALESCE(p.commission, 0) / r.revenue
                       END AS commission_share,
                       CASE WHEN r.revenue = 0 THEN 0
                            ELSE -COALESCE(p.logistics, 0) / r.revenue
                       END AS logistics_share,
                       CASE WHEN r.revenue = 0 THEN 0
                            ELSE COALESCE(p.points, 0) / r.revenue
                       END AS points_share
                FROM runs AS r
                LEFT JOIN (
                    SELECT run_id,
                           SUM(commission) AS commission,
                           SUM(
                               processing + delivery + logistics
                               + reverse_logistics + returns_cancels
                           ) AS logistics,
                           SUM(points) AS points
                    FROM product_results
                    GROUP BY run_id
                ) AS p ON p.run_id = r.id
                ORDER BY
                    COALESCE(r.period_start, r.period_end, substr(r.created_at, 1, 10)) ASC,
                    COALESCE(r.period_end, r.period_start, substr(r.created_at, 1, 10)) ASC,
                    r.id ASC
                """
            ).fetchall()
        summaries: list[RunSummary] = []
        for row in rows:
            values = dict(row)
            unallocated_income_tax = float(values.pop("unallocated_income_tax"))
            revenue = float(values["revenue"])
            values["net_margin"] = (
                (
                    float(values["net_profit"])
                    + float(values["unallocated_total"])
                    - unallocated_income_tax
                )
                / revenue
                if revenue
                else 0.0
            )
            summaries.append(RunSummary(**values))
        return summaries

    def rename_run(self, run_id: int, report_name: str) -> None:
        cleaned = " ".join(str(report_name).split())
        if not cleaned:
            raise ValueError("Наименование отчета не может быть пустым")
        if len(cleaned) > 200:
            raise ValueError("Наименование отчета не должно превышать 200 символов")
        with self.transaction() as db:
            cursor = db.execute(
                "UPDATE runs SET report_name = ? WHERE id = ?",
                (cleaned, run_id),
            )
            if cursor.rowcount == 0:
                raise KeyError(f"Расчет #{run_id} не найден")

    def set_run_created_at(self, run_id: int, created_at: str) -> None:
        with self.transaction() as db:
            cursor = db.execute(
                "UPDATE runs SET created_at = ? WHERE id = ?",
                (created_at, run_id),
            )
            if cursor.rowcount == 0:
                raise KeyError(f"Расчет #{run_id} не найден")

    def delete_run(self, run_id: int) -> int:
        with self.transaction() as db:
            stored_paths = [
                str(row["stored_path"])
                for row in db.execute(
                    "SELECT stored_path FROM source_files WHERE run_id = ?",
                    (run_id,),
                ).fetchall()
            ]
            cursor = db.execute("DELETE FROM runs WHERE id = ?", (run_id,))
            if cursor.rowcount == 0:
                raise KeyError(f"Расчет #{run_id} не найден")

        return self._remove_unreferenced_source_files(stored_paths)

    def _remove_unreferenced_source_files(self, stored_paths: list[str]) -> int:
        if not stored_paths:
            return 0
        with self.read() as db:
            remaining_paths = {
                _resolved_path(str(row["stored_path"]))
                for row in db.execute("SELECT DISTINCT stored_path FROM source_files").fetchall()
            }

        source_root = (self.path.parent / "source_files").resolve()
        removed_files = 0
        for stored_path in set(stored_paths):
            candidate = _resolved_path(stored_path)
            if candidate in remaining_paths or not _is_relative_to(candidate, source_root):
                continue
            try:
                existed = candidate.is_file()
                candidate.unlink(missing_ok=True)
                removed_files += int(existed)
            except OSError:
                # Запись истории уже удалена. Недоступную файловую
                # копию можно безопасно очистить позже.
                continue
        return removed_files

    def load_calculation(self, run_id: int) -> RunCalculation:
        with self.read() as db:
            run = db.execute("SELECT * FROM runs WHERE id = ?", (run_id,)).fetchone()
            if run is None:
                raise KeyError(f"Расчет #{run_id} не найден")
            result_rows = db.execute(
                "SELECT * FROM product_results WHERE run_id = ? ORDER BY article COLLATE NOCASE",
                (run_id,),
            ).fetchall()
            breakdown_rows = db.execute(
                "SELECT accrual_type, row_count, amount FROM unallocated WHERE run_id = ? ORDER BY accrual_type COLLATE NOCASE",
                (run_id,),
            ).fetchall()
            stat_rows = db.execute(
                "SELECT accrual_type, with_article, without_article FROM accrual_stats WHERE run_id = ? ORDER BY accrual_type COLLATE NOCASE",
                (run_id,),
            ).fetchall()
            events = db.execute(
                "SELECT event_type, message FROM quality_events WHERE run_id = ?",
                (run_id,),
            ).fetchall()
            tax_schedule = _read_run_tax_schedule(db, run_id)
        skipped = {row["message"].split(" — ", 1)[0]: row["message"] for row in events if row["event_type"] == "Пропущенный артикул"}
        conflicts = {
            row["message"].removeprefix("SKU ").removesuffix(" связан с несколькими артикулами")
            for row in events
            if row["event_type"] == "Конфликт SKU"
        }
        source_period_warnings = [
            str(row["message"])
            for row in events
            if row["event_type"] == "Несовпадение периодов"
        ]
        double_count_warnings = [
            str(row["message"])
            for row in events
            if row["event_type"] == DOUBLE_COUNT_EVENT
        ]
        tax_rate_warnings = [
            str(row["message"])
            for row in events
            if row["event_type"] == TAX_RATE_EVENT
        ]
        tax_period_start = _parse_date(run["tax_period_start"])
        tax_period_end = _parse_date(run["tax_period_end"])
        return RunCalculation(
            run_id=run_id,
            period_start=_parse_date(run["period_start"]),
            period_end=_parse_date(run["period_end"]),
            tax_rate=float(run["tax_rate"]),
            products=[_row_to_product_result(row) for row in result_rows],
            unallocated_total=float(run["unallocated_total"]),
            unallocated={row["accrual_type"]: (int(row["row_count"]), float(row["amount"])) for row in breakdown_rows},
            accrual_stats={row["accrual_type"]: (int(row["with_article"]), int(row["without_article"])) for row in stat_rows},
            skipped_articles=skipped,
            sku_conflicts=conflicts,
            duplicate_realization_rows=int(run["duplicate_realization_rows"]),
            already_accrued_realization_rows=int(run["already_accrued_realization_rows"]),
            realization_revenue=float(run["realization_revenue"]),
            realization_units=float(run["realization_units"]),
            source_period_warnings=source_period_warnings,
            taxable_unallocated_income_override=float(
                run["taxable_unallocated_income"]
            ),
            double_count_warnings=double_count_warnings,
            unallocated_income_tax_override=float(run["unallocated_income_tax"]),
            tax_schedule=tax_schedule,
            tax_period_start=tax_period_start,
            tax_period_end=tax_period_end,
            applied_tax_rates=(
                tax_schedule.applied_rates(tax_period_start, tax_period_end)
                if tax_period_start is not None
                else []
            ),
            tax_rate_warnings=tax_rate_warnings,
        )

    def list_source_files(self, run_id: int) -> list[dict[str, object]]:
        with self.read() as db:
            rows = db.execute(
                "SELECT * FROM source_files WHERE run_id = ? ORDER BY id",
                (run_id,),
            ).fetchall()
        return [dict(row) for row in rows]

    def list_quality_events(self, run_id: int) -> list[dict[str, object]]:
        with self.read() as db:
            rows = db.execute(
                "SELECT severity, event_type, message FROM quality_events WHERE run_id = ? ORDER BY id",
                (run_id,),
            ).fetchall()
        return [dict(row) for row in rows]

    def accrual_guide(self, run_id: int) -> list[dict[str, object]]:
        with self.read() as db:
            current_rows = db.execute(
                "SELECT normalized_type, accrual_type, with_article, without_article FROM accrual_stats WHERE run_id = ?",
                (run_id,),
            ).fetchall()
            history_rows = db.execute(
                """
                SELECT normalized_type, MIN(accrual_type) AS accrual_type,
                       SUM(with_article) AS with_article, SUM(without_article) AS without_article
                FROM accrual_stats GROUP BY normalized_type
                """
            ).fetchall()
        current = {row["normalized_type"]: row for row in current_rows}
        history = {row["normalized_type"]: row for row in history_rows}
        items: list[dict[str, object]] = []
        for key in sorted(set(current) | set(history)):
            current_row = current.get(key)
            history_row = history.get(key)
            name = (current_row or history_row)["accrual_type"]
            current_with = int(current_row["with_article"]) if current_row else 0
            current_without = int(current_row["without_article"]) if current_row else 0
            history_with = int(history_row["with_article"]) if history_row else 0
            history_without = int(history_row["without_article"]) if history_row else 0
            items.append(
                {
                    "accrual_type": name,
                    "category": guide_target(name),
                    "current_status": distribution_status(current_with, current_without, "НЕ В ТЕКУЩЕМ ЗАПУСКЕ"),
                    "current_with": current_with,
                    "current_without": current_without,
                    "history_status": distribution_status(history_with, history_without, "НЕТ ИСТОРИИ"),
                    "history_with": history_with,
                    "history_without": history_without,
                }
            )
        return items

    def save_planned_price(self, run_id: int, article: str, value: float) -> None:
        with self.transaction() as db:
            db.execute(
                """
                INSERT INTO scenario_prices(run_id, article, planned_price) VALUES (?, ?, ?)
                ON CONFLICT(run_id, article) DO UPDATE SET
                    planned_price = excluded.planned_price,
                    updated_at = CURRENT_TIMESTAMP
                """,
                (run_id, article, value),
            )

    def planned_prices(self, run_id: int) -> dict[str, float]:
        with self.read() as db:
            rows = db.execute(
                "SELECT article, planned_price FROM scenario_prices WHERE run_id = ?",
                (run_id,),
            ).fetchall()
        return {row["article"]: float(row["planned_price"]) for row in rows}

    def clear_planned_prices(self, run_id: int) -> None:
        with self.transaction() as db:
            db.execute("DELETE FROM scenario_prices WHERE run_id = ?", (run_id,))


def _date_text(value: date | None) -> str | None:
    return value.isoformat() if value else None


def _parse_date(value: str | None) -> date | None:
    return date.fromisoformat(value) if value else None


def _default_run_name(run_id: int, period_start: str | None, period_end: str | None) -> str:
    start = _display_date(period_start)
    end = _display_date(period_end)
    if start and end and start != end:
        return f"Отчет Ozon за {start}–{end}"
    if start or end:
        return f"Отчет Ozon за {start or end}"
    return "Отчет Ozon без периода"


def _display_date(value: str | None) -> str:
    if not value:
        return ""
    try:
        return date.fromisoformat(value).strftime("%d.%m.%Y")
    except ValueError:
        return value


def _resolved_path(value: str) -> Path:
    return Path(value).expanduser().resolve()


def _is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _product_result_tuple(run_id: int, item: ProductResult) -> tuple[object, ...]:
    return (
        run_id,
        item.article,
        item.name,
        item.category,
        item.material_cost,
        item.labor_cost,
        item.units,
        item.revenue_no_points,
        item.partner_programs,
        item.points,
        item.commission,
        item.processing,
        item.delivery,
        item.logistics,
        item.reverse_logistics,
        item.returns_cancels,
        item.acquiring,
        item.stars,
        item.packaging,
        item.compensation,
        item.other,
        item.financial_result,
    )


def _row_to_product_result(row: sqlite3.Row) -> ProductResult:
    return ProductResult(
        article=row["article"],
        name=row["name"],
        category=str(row["category"] or ""),
        material_cost=float(row["material_cost"]),
        labor_cost=float(row["labor_cost"]),
        units=float(row["units"]),
        revenue_no_points=float(row["revenue_no_points"]),
        partner_programs=float(row["partner_programs"]),
        points=float(row["points"]),
        commission=float(row["commission"]),
        processing=float(row["processing"]),
        delivery=float(row["delivery"]),
        logistics=float(row["logistics"]),
        reverse_logistics=float(row["reverse_logistics"]),
        returns_cancels=float(row["returns_cancels"]),
        acquiring=float(row["acquiring"]),
        stars=float(row["stars"]),
        packaging=float(row["packaging"]),
        compensation=float(row["compensation"]),
        other=float(row["other"]),
        financial_result=float(row["financial_result"]),
        tax_override=float(row["tax"]),
    )


def _read_tax_schedule(db: sqlite3.Connection) -> TaxRateSchedule:
    rows = db.execute("SELECT valid_from, rate FROM tax_rates").fetchall()
    return TaxRateSchedule(
        TaxRatePeriod(_parse_date(row["valid_from"]), float(row["rate"])) for row in rows
    )


def _read_run_tax_schedule(db: sqlite3.Connection, run_id: int) -> TaxRateSchedule:
    rows = db.execute(
        "SELECT valid_from, rate FROM run_tax_rates WHERE run_id = ? ORDER BY position",
        (run_id,),
    ).fetchall()
    if not rows:
        run = db.execute("SELECT tax_rate FROM runs WHERE id = ?", (run_id,)).fetchone()
        return TaxRateSchedule.single(float(run["tax_rate"]) if run else DEFAULT_TAX_RATE)
    return TaxRateSchedule(
        TaxRatePeriod(_parse_date(row["valid_from"]), float(row["rate"])) for row in rows
    )
