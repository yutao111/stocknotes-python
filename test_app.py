import io
import gc
import json
import os
import re
import sqlite3
import tempfile
import threading
import time
import unittest
from datetime import date, datetime, timedelta
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook


class StockNotesTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.TemporaryDirectory()
        os.environ["STOCKNOTES_DB"] = os.path.join(cls.temp_dir.name, "test.db")
        import app
        cls.module = app
        cls.client = app.app.test_client()

    @classmethod
    def tearDownClass(cls):
        gc.collect()
        cls.temp_dir.cleanup()

    def setUp(self):
        self.module.HOT_SECTORS_CACHE.update({"modules": {}, "updated_at": None})
        with self.module.db_connect() as db:
            db.execute("DELETE FROM alert_signal_outcomes")
            db.execute("DELETE FROM alert_signal_samples")
            db.execute("DELETE FROM notifications")
            db.execute("DELETE FROM alert_rule_states")
            db.execute(
                "UPDATE alert_types SET params_json = ?, enabled = 1 WHERE code = ?",
                (json.dumps(self.module.THREE_DAY_DIP_DEFAULT_PARAMS), self.module.THREE_DAY_DIP_CODE),
            )
            db.execute(
                "UPDATE alert_types SET params_json = ?, enabled = 1 WHERE code = ?",
                (json.dumps(self.module.INTRADAY_REBOUND_DEFAULT_PARAMS), self.module.INTRADAY_REBOUND_CODE),
            )
            db.execute("DELETE FROM account_snapshots")
            db.execute("DELETE FROM watchlist_stocks")
            db.execute("DELETE FROM trade_excursion_metrics")
            db.execute("DELETE FROM daily_prices")
            db.execute("DELETE FROM intraday_quotes")
            db.execute("DELETE FROM trade_review_reason_types")
            db.execute("DELETE FROM trade_reviews")
            db.execute("DELETE FROM trade_episode_executions")
            db.execute("DELETE FROM trade_episodes")
            db.execute("DELETE FROM fifo_matches")
            db.execute("DELETE FROM positions")
            db.execute("DELETE FROM unmatched_sells")
            db.execute("DELETE FROM executions")
            db.execute("DELETE FROM import_jobs")
            db.execute("DELETE FROM users WHERE name <> 'yutaoGS'")
            user = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()
        with self.client.session_transaction() as session:
            session["user_id"] = user["id"]

    def make_statement(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["成交日期", "成交时间", "证券代码", "证券名称", "操作", "成交数量", "成交编号", "成交价格", "成交金额", "佣金", "印花税", "其他杂费", "其他费", "过户费", "交易市场"])
        sheet.append([20260801, "09:31:00", 2156, "通富微电", "证券买入", 100, "B1", 10, 1000, 1, 0, 0, 0, 0, "深圳A股"])
        sheet.append([20260802, "10:00:00", 2156, "通富微电", "证券买入", 100, "B2", 12, 1200, 1, 0, 0, 0, 0, "深圳A股"])
        sheet.append([20260803, "13:00:00", 2156, "通富微电", "证券卖出", -150, "S1", 14, 2100, 1, 2, 0, 0, 0, "深圳A股"])
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return output

    def import_statement(self, days=1):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["成交日期", "成交时间", "证券代码", "证券名称", "操作", "成交数量", "成交编号", "成交价格", "成交金额", "佣金", "印花税", "其他杂费", "其他费", "过户费", "交易市场"])
        for day in range(1, days + 1):
            sheet.append([20260100 + day, f"{9 + day % 9}:00:00", 2156, "通富微电", "证券买入", 100, f"B{day}", 10, 1000, 1, 0, 0, 0, 0, "深圳A股"])
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        response = self.client.post("/admin/preview", data={"statement": (output, "statement.xlsx")}, content_type="multipart/form-data")
        self.assertEqual(response.status_code, 200)
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW' ORDER BY created_at DESC, rowid DESC LIMIT 1").fetchone()
        response = self.client.post(f"/admin/import/{job['id']}", follow_redirects=True)
        self.assertEqual(response.status_code, 200)

    def test_import_fifo_and_duplicate(self):
        response = self.client.post("/admin/preview", data={"statement": (self.make_statement(), "statement.xlsx")}, content_type="multipart/form-data")
        self.assertEqual(response.status_code, 200)
        preview_text = response.get_data(as_text=True)
        self.assertIn("解析完成", preview_text)
        self.assertIn("有效记录", preview_text)
        self.assertIn("预计新增", preview_text)
        self.assertIn("确认导入 3 条记录", preview_text)
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs ORDER BY created_at DESC LIMIT 1").fetchone()
        response = self.client.post(f"/admin/import/{job['id']}", follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM executions").fetchone()[0], 3)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM fifo_matches").fetchone()[0], 2)
            position = db.execute("SELECT * FROM positions").fetchone()
            self.assertEqual(position["stock_code"], "002156")
            self.assertAlmostEqual(position["quantity"], 50)
            profit = db.execute("SELECT SUM(profit) FROM fifo_matches").fetchone()[0]
            self.assertAlmostEqual(profit, 495.5, places=2)
            episode = db.execute("SELECT * FROM trade_episodes").fetchone()
            self.assertEqual(episode["status"], "OPEN")
            self.assertEqual(db.execute("SELECT COUNT(*) FROM trade_episode_executions").fetchone()[0], 3)

        response = self.client.post("/admin/preview", data={"statement": (self.make_statement(), "statement.xlsx")}, content_type="multipart/form-data")
        self.assertEqual(response.status_code, 200)
        with self.module.db_connect() as db:
            duplicate_job = db.execute("SELECT * FROM import_jobs WHERE status = 'PREVIEW' ORDER BY created_at DESC, rowid DESC LIMIT 1").fetchone()
            self.assertEqual(duplicate_job["duplicate_rows"], 3)
        self.client.post(f"/admin/import/{duplicate_job['id']}")
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM executions").fetchone()[0], 3)

    def test_manual_execution_is_reconciled_by_statement(self):
        response = self.client.post(
            "/admin/manual",
            data={
                "trade_date": "2026-08-01",
                "trade_time": "09:30:00",
                "stock_code": "2156",
                "stock_name": "通富微电",
                "action": "BUY",
                "quantity": "100",
                "deal_price": "10",
                "commission": "0",
                "stamp_tax": "0",
                "transfer_fee": "0",
            },
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn("已补录当天成交", response.get_data(as_text=True))
        with self.module.db_connect() as db:
            manual = db.execute("SELECT * FROM executions").fetchone()
            self.assertEqual(manual["source"], "MANUAL")
            self.assertEqual(db.execute("SELECT COUNT(*) FROM positions").fetchone()[0], 1)

        self.client.post(
            "/admin/preview",
            data={"statement": (self.make_statement(), "statement.xlsx")},
            content_type="multipart/form-data",
        )
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW' ORDER BY created_at DESC, rowid DESC LIMIT 1").fetchone()
        response = self.client.post(f"/admin/import/{job['id']}", follow_redirects=True)
        text = response.get_data(as_text=True)
        self.assertIn("转正补录 1 条", text)
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM executions").fetchone()[0], 3)
            reconciled = db.execute("SELECT * FROM executions WHERE deal_id = 'B1'").fetchone()
            self.assertEqual(reconciled["source"], "STATEMENT")
            self.assertEqual(reconciled["trade_time"], "09:31:00")
            self.assertAlmostEqual(reconciled["commission"], 1)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM fifo_matches").fetchone()[0], 2)

    def test_manual_execution_validation(self):
        response = self.client.post(
            "/admin/manual",
            data={
                "trade_date": "2026-08-01",
                "trade_time": "09:30:00",
                "stock_code": "2156",
                "stock_name": "通富微电",
                "action": "BUY",
                "quantity": "0",
                "deal_price": "10",
                "commission": "0",
                "stamp_tax": "0",
                "transfer_fee": "0",
            },
            follow_redirects=True,
        )
        self.assertIn("成交数量必须大于 0", response.get_data(as_text=True))
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM executions").fetchone()[0], 0)

    def test_trade_episode_groups_multiple_buys_and_split_sells(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["成交日期", "成交时间", "证券代码", "证券名称", "操作", "成交数量", "成交编号", "成交价格", "成交金额"])
        sheet.append([20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10, 1000])
        sheet.append([20260802, "10:00:00", 2156, "通富微电", "证券买入", 100, "B2", 12, 1200])
        sheet.append([20260803, "11:00:00", 2156, "通富微电", "证券卖出", -50, "S1", 13, 650])
        sheet.append([20260804, "14:00:00", 2156, "通富微电", "证券卖出", -150, "S2", 14, 2100])
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        self.client.post("/admin/preview", data={"statement": (output, "episode.xlsx")}, content_type="multipart/form-data")
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW'").fetchone()
        self.client.post(f"/admin/import/{job['id']}")

        with self.module.db_connect() as db:
            episodes = db.execute("SELECT * FROM trade_episodes").fetchall()
            self.assertEqual(len(episodes), 1)
            episode_id = episodes[0]["id"]
            self.assertEqual(episodes[0]["status"], "CLOSED")
            links = db.execute("SELECT role FROM trade_episode_executions ORDER BY execution_id").fetchall()
            self.assertEqual([row["role"] for row in links], ["ENTRY", "ADD", "REDUCE", "EXIT"])

        detail = self.client.get(f"/analysis/trades/{episode_id}")
        self.assertEqual(detail.status_code, 200)
        text = detail.get_data(as_text=True)
        self.assertIn("交易复盘详情", text)
        self.assertIn("通富微电", text)
        self.assertIn("首次建仓", text)
        self.assertIn("完全退出", text)
        self.assertIn("我的交易复盘", text)

        self.client.post("/admin/rebuild")
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT id FROM trade_episodes").fetchone()["id"], episode_id)
        self.assertEqual(self.client.get(f"/analysis/trades/{episode_id}").status_code, 200)

    def test_full_exit_then_reentry_creates_new_episode(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["成交日期", "成交时间", "证券代码", "证券名称", "操作", "成交数量", "成交编号", "成交价格"])
        sheet.append([20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10])
        sheet.append([20260802, "10:00:00", 2156, "通富微电", "证券卖出", -100, "S1", 11])
        sheet.append([20260803, "11:00:00", 2156, "通富微电", "证券买入", 200, "B2", 12])
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        self.client.post("/admin/preview", data={"statement": (output, "reentry.xlsx")}, content_type="multipart/form-data")
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW'").fetchone()
        self.client.post(f"/admin/import/{job['id']}")
        with self.module.db_connect() as db:
            episodes = db.execute("SELECT status FROM trade_episodes ORDER BY opened_at").fetchall()
            self.assertEqual([row["status"] for row in episodes], ["CLOSED", "OPEN"])

    def _insert_review(self, episode_id, user_id=None, **overrides):
        with self.module.db_connect() as db:
            user_id = user_id or db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            episode = db.execute(
                "SELECT * FROM trade_episodes WHERE id = ? AND user_id = ?", (episode_id, user_id)
            ).fetchone()
            fields = {
                "user_id": user_id,
                "trade_episode_id": episode_id,
                "review_status": "PENDING",
                "trade_reason": "回调买入",
                "expected_profit_percent": 20.0,
                "expected_target_price": 12.0,
                "stop_loss_price": 9.0,
                "expected_holding_days": 10,
                "confidence_level": 4,
                "sell_reason": "达到目标",
                "judgement_result": "CORRECT",
                "main_problem": "NO_OBVIOUS_PROBLEM",
                "review_note": "执行顺利",
                "next_action": "继续按计划执行",
                "original_opening_execution_id": episode["opening_execution_id"],
                "stock_code_snapshot": episode["stock_code"],
                "opened_at_snapshot": episode["opened_at"],
                "closed_at_snapshot": episode["closed_at"],
                "created_at": "2026-08-21T00:00:00",
                "updated_at": "2026-08-21T00:00:00",
                "completed_at": None,
            }
            fields.update(overrides)
            columns = ", ".join(fields)
            placeholders = ", ".join("?" for _ in fields)
            db.execute(
                f"INSERT INTO trade_reviews ({columns}) VALUES ({placeholders})",
                tuple(fields.values()),
            )

    def test_trade_review_schema_and_constraints(self):
        self.client.post("/admin/preview", data={"statement": (self.make_statement(), "s.xlsx")}, content_type="multipart/form-data")
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW'").fetchone()
            tables = {row["name"] for row in db.execute("SELECT name FROM sqlite_master WHERE type = 'table'")}
        self.client.post(f"/admin/import/{job['id']}")
        self.assertIn("trade_reviews", tables)
        self.assertIn("trade_review_reason_types", tables)

        with self.module.db_connect() as db:
            episode_id = db.execute("SELECT id FROM trade_episodes").fetchone()["id"]
            user_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
        self._insert_review(episode_id)
        with self.module.db_connect() as db:
            review_id = db.execute("SELECT id FROM trade_reviews").fetchone()["id"]
        with self.module.db_connect() as db:
            db.execute("INSERT INTO trade_review_reason_types (trade_review_id, reason_type) VALUES (?, 'PULLBACK')", (review_id,))
            db.execute("INSERT INTO trade_review_reason_types (trade_review_id, reason_type) VALUES (?, 'TECHNICAL_PATTERN')", (review_id,))
            self.assertEqual(db.execute("SELECT COUNT(*) FROM trade_review_reason_types").fetchone()[0], 2)
        with self.module.db_connect() as db:
            with self.assertRaises(sqlite3.IntegrityError):
                db.execute("INSERT INTO trade_review_reason_types (trade_review_id, reason_type) VALUES (?, 'PULLBACK')", (review_id,))
            with self.assertRaises(sqlite3.IntegrityError):
                db.execute("INSERT INTO trade_review_reason_types (trade_review_id, reason_type) VALUES (?, 'INVALID')", (review_id,))
        with self.assertRaises(sqlite3.IntegrityError):
            self._insert_review(episode_id, user_id=user_id)
        with self.assertRaises(sqlite3.IntegrityError):
            self._insert_review(episode_id, confidence_level=0)
        with self.assertRaises(sqlite3.IntegrityError):
            self._insert_review(episode_id, confidence_level=6)
        with self.assertRaises(sqlite3.IntegrityError):
            self._insert_review(episode_id, expected_target_price=0)
        with self.assertRaises(sqlite3.IntegrityError):
            self._insert_review(episode_id, main_problem="INVALID")
        with self.assertRaises(sqlite3.IntegrityError):
            self._insert_review(episode_id, review_status="INVALID")

        self.client.post("/users/create", data={"name": "second", "next": "/analysis/stocks"})
        with self.module.db_connect() as db:
            second_id = db.execute("SELECT id FROM users WHERE name = 'second'").fetchone()["id"]
        self.client.post("/users/switch", data={"user_id": second_id, "next": "/analysis/stocks"})
        self.client.post("/admin/preview", data={"statement": (self.make_statement(), "second.xlsx")}, content_type="multipart/form-data")
        with self.module.db_connect() as db:
            second_job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW'").fetchone()
        self.client.post(f"/admin/import/{second_job['id']}")
        with self.module.db_connect() as db:
            second_episode = db.execute("SELECT id FROM trade_episodes WHERE user_id = ?", (second_id,)).fetchone()["id"]
        self._insert_review(second_episode, user_id=second_id)
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM trade_reviews").fetchone()[0], 2)
            db.execute("DELETE FROM users WHERE id = ?", (second_id,))
            self.assertEqual(db.execute("SELECT COUNT(*) FROM trade_reviews").fetchone()[0], 1)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM trade_review_reason_types").fetchone()[0], 2)

    def test_trade_review_survives_episode_replacement(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["成交日期", "成交时间", "证券代码", "证券名称", "操作", "成交数量", "成交编号", "成交价格"])
        sheet.append([20260810, "09:30:00", 2156, "通富微电", "证券买入", 100, "B10", 10])
        sheet.append([20260811, "10:00:00", 2156, "通富微电", "证券卖出", -100, "S11", 11])
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        self.client.post("/admin/preview", data={"statement": (output, "first.xlsx")}, content_type="multipart/form-data")
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW'").fetchone()
        self.client.post(f"/admin/import/{job['id']}")
        with self.module.db_connect() as db:
            episode_id = db.execute("SELECT id FROM trade_episodes").fetchone()["id"]
            opening_id = db.execute("SELECT opening_execution_id FROM trade_episodes").fetchone()["opening_execution_id"]
            opened_at = db.execute("SELECT opened_at FROM trade_episodes").fetchone()["opened_at"]
        self._insert_review(episode_id)

        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["成交日期", "成交时间", "证券代码", "证券名称", "操作", "成交数量", "成交编号", "成交价格"])
        sheet.append([20260801, "09:30:00", 2156, "通富微电", "证券买入", 200, "B01", 8])
        sheet.append([20260803, "10:00:00", 2156, "通富微电", "证券卖出", -100, "S03", 9])
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        self.client.post("/admin/preview", data={"statement": (output, "earlier.xlsx")}, content_type="multipart/form-data")
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW'").fetchone()
        self.client.post(f"/admin/import/{job['id']}")

        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM trade_episodes").fetchone()[0], 1)
            review = db.execute("SELECT * FROM trade_reviews").fetchone()
            self.assertIsNotNone(review)
            self.assertIsNone(review["trade_episode_id"])
            self.assertEqual(review["original_opening_execution_id"], opening_id)
            self.assertEqual(review["stock_code_snapshot"], "002156")
            self.assertEqual(review["opened_at_snapshot"], opened_at)
            self.assertEqual(review["review_note"], "执行顺利")

    def _import_workbook_rows(self, rows):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["成交日期", "成交时间", "证券代码", "证券名称", "操作", "成交数量", "成交编号", "成交价格"])
        for row in rows:
            sheet.append(row)
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        self.client.post("/admin/preview", data={"statement": (output, "ep.xlsx")}, content_type="multipart/form-data")
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW' ORDER BY created_at DESC, rowid DESC LIMIT 1").fetchone()
        self.client.post(f"/admin/import/{job['id']}")
        with self.module.db_connect() as db:
            return db.execute("SELECT id FROM trade_episodes ORDER BY id DESC LIMIT 1").fetchone()["id"]

    def test_trade_review_form_renders_and_saves(self):
        episode_id = self._import_workbook_rows([
            [20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
            [20260802, "10:00:00", 2156, "通富微电", "证券卖出", -100, "S1", 12],
        ])
        page = self.client.get(f"/analysis/trades/{episode_id}").get_data(as_text=True)
        self.assertIn("我的交易复盘", page)
        self.assertIn("未复盘", page)
        self.assertIn("为什么买？", page)
        self.assertIn("回调买入", page)
        self.assertIn("标记已复盘", page)

        response = self.client.post(
            f"/analysis/trades/{episode_id}/review",
            data={
                "reason_type": ["PULLBACK", "TECHNICAL_PATTERN"],
                "trade_reason": "回调到支撑位",
                "sell_reason": "达到目标价",
                "expected_profit_percent": "20",
                "expected_target_price": "12",
                "stop_loss_price": "9",
                "expected_holding_days": "10",
                "confidence_level": "4",
                "judgement_result": "CORRECT",
                "main_problem": "NO_OBVIOUS_PROBLEM",
                "review_note": "按计划执行",
                "next_action": "保持节奏",
                "action": "draft",
                "sort": "recent",
            },
            follow_redirects=True,
        )
        self.assertIn("复盘草稿已保存", response.get_data(as_text=True))
        with self.module.db_connect() as db:
            review = db.execute("SELECT * FROM trade_reviews").fetchone()
            self.assertEqual(review["review_status"], "PENDING")
            self.assertEqual(review["trade_reason"], "回调到支撑位")
            self.assertAlmostEqual(review["expected_target_price"], 12)
            self.assertEqual(review["confidence_level"], 4)
            reason_types = {row["reason_type"] for row in db.execute("SELECT reason_type FROM trade_review_reason_types")}
            self.assertEqual(reason_types, {"PULLBACK", "TECHNICAL_PATTERN"})

        page = self.client.get(f"/analysis/trades/{episode_id}").get_data(as_text=True)
        self.assertIn("草稿", page)
        self.assertIn('name="reason_type" value="PULLBACK" checked', page)
        self.assertIn('value="12.0"', page)

        response = self.client.post(
            f"/analysis/trades/{episode_id}/review",
            data={"action": "complete", "trade_reason": "x", "judgement_result": "CORRECT", "sort": "recent"},
            follow_redirects=True,
        )
        self.assertIn("复盘已完成并保存", response.get_data(as_text=True))
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT review_status FROM trade_reviews").fetchone()["review_status"], "COMPLETED")
            self.assertIsNotNone(db.execute("SELECT completed_at FROM trade_reviews").fetchone()["completed_at"])

    def test_trade_review_open_episode_only_draft_and_validation(self):
        episode_id = self._import_workbook_rows([
            [20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
        ])
        response = self.client.post(
            f"/analysis/trades/{episode_id}/review",
            data={"action": "complete", "trade_reason": "买入持有", "sort": "recent"},
            follow_redirects=True,
        )
        self.assertIn("复盘草稿已保存", response.get_data(as_text=True))
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT review_status FROM trade_reviews").fetchone()["review_status"], "PENDING")
            self.assertIsNone(db.execute("SELECT completed_at FROM trade_reviews").fetchone()["completed_at"])
        response = self.client.post(
            f"/analysis/trades/{episode_id}/review",
            data={"action": "complete", "trade_reason": "买入持有", "judgement_result": "CORRECT", "sort": "recent"},
            follow_redirects=True,
        )
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT review_status FROM trade_reviews").fetchone()["review_status"], "PENDING")

        response = self.client.post(
            f"/analysis/trades/{episode_id}/review",
            data={"action": "complete", "trade_reason": "x", "judgement_result": "INVALID", "sort": "recent"},
            follow_redirects=True,
        )
        self.assertIn("判断结果无效", response.get_data(as_text=True))

    def test_trade_review_save_requires_judgement_for_complete(self):
        episode_id = self._import_workbook_rows([
            [20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
            [20260802, "10:00:00", 2156, "通富微电", "证券卖出", -100, "S1", 12],
        ])
        response = self.client.post(
            f"/analysis/trades/{episode_id}/review",
            data={"action": "complete", "trade_reason": "x", "sort": "recent"},
            follow_redirects=True,
        )
        self.assertIn("标记已复盘前请选择判断结果", response.get_data(as_text=True))
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM trade_reviews").fetchone()[0], 0)

    def test_trade_review_save_is_user_isolated(self):
        episode_id = self._import_workbook_rows([
            [20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
            [20260802, "10:00:00", 2156, "通富微电", "证券卖出", -100, "S1", 12],
        ])
        self.client.post("/users/create", data={"name": "second", "next": "/analysis/stocks"})
        response = self.client.post(
            f"/analysis/trades/{episode_id}/review",
            data={"action": "draft", "trade_reason": "x", "sort": "recent"},
        )
        self.assertEqual(response.status_code, 404)
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM trade_reviews").fetchone()[0], 0)

    def test_review_progress_and_reviews_page(self):
        first = self._import_workbook_rows([
            [20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
            [20260802, "10:00:00", 2156, "通富微电", "证券卖出", -100, "S1", 12],
        ])
        second = self._import_workbook_rows([
            [20260803, "09:30:00", 100002, "国机精工", "证券买入", 200, "B2", 8],
            [20260804, "10:00:00", 100002, "国机精工", "证券卖出", -200, "S2", 9],
        ])
        self._import_workbook_rows([
            [20260805, "09:30:00", 300540, "蜀道装备", "证券买入", 100, "B3", 15],
        ])
        with self.module.db_connect() as db:
            user_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            progress = self.module.review_progress_data(db, user_id)
            self.assertEqual(progress["total"], 2)
            self.assertEqual(progress["completed"], 0)
            self.assertEqual(progress["pending"], 2)
            self.assertAlmostEqual(progress["rate"], 0.0)

        pending = self.client.get("/analysis/reviews?status=pending&start_date=2026-08-01&end_date=2026-08-31").get_data(as_text=True)
        self.assertIn("待复盘 2", pending)
        self.assertIn("去复盘 →", pending)
        self.assertIn("通富微电", pending)
        self.assertIn("国机精工", pending)
        self.assertNotIn("蜀道装备", pending)
        completed_page = self.client.get("/analysis/reviews?status=completed&start_date=2026-08-01&end_date=2026-08-31").get_data(as_text=True)
        self.assertIn("还没有已完成复盘的交易", completed_page)

        self.client.post(
            f"/analysis/trades/{first}/review",
            data={
                "action": "complete",
                "trade_reason": "x",
                "judgement_result": "CORRECT",
                "reason_type": ["PULLBACK", "TECHNICAL_PATTERN"],
                "sort": "recent",
            },
            follow_redirects=True,
        )
        with self.module.db_connect() as db:
            progress = self.module.review_progress_data(db, user_id)
            self.assertEqual(progress["completed"], 1)
            self.assertEqual(progress["pending"], 1)
            self.assertAlmostEqual(progress["rate"], 0.5)

        pending = self.client.get("/analysis/reviews?status=pending&start_date=2026-08-01&end_date=2026-08-31").get_data(as_text=True)
        self.assertIn("待复盘 1", pending)
        self.assertIn("国机精工", pending)
        self.assertNotIn("通富微电", pending)
        completed_page = self.client.get("/analysis/reviews?status=completed&start_date=2026-08-01&end_date=2026-08-31").get_data(as_text=True)
        self.assertIn("已复盘 1", completed_page)
        self.assertIn("通富微电", completed_page)
        self.assertIn("去复盘 →", completed_page)
        self.assertIn('id="review-modal-backdrop"', completed_page)
        self.assertIn("review-popup-trigger", completed_page)
        self.assertIn("CORRECT", completed_page)
        match = re.search(r"data-review='([^']+)'", completed_page)
        self.assertIsNotNone(match)
        review_data = json.loads(match.group(1))
        self.assertEqual(review_data["code"], "002156")
        self.assertEqual(review_data["judgement"], "CORRECT")
        self.assertEqual(review_data["trade_reason"], "x")
        self.assertEqual(review_data["reason_types"], ["回调买入", "技术形态"])

        summary = self.client.get("/analysis/summary").get_data(as_text=True)
        self.assertIn("交易复盘完成度", summary)
        self.assertIn("待复盘交易", summary)

    def test_review_date_filter(self):
        today = date.today()
        self._import_workbook_rows([
            [(today - timedelta(days=1)).strftime("%Y%m%d"), "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
            [today.strftime("%Y%m%d"), "10:00:00", 2156, "通富微电", "证券卖出", -100, "S1", 12],
        ])
        self._import_workbook_rows([
            ["20260102", "09:30:00", 100002, "国机精工", "证券买入", 100, "B2", 8],
            ["20260103", "10:00:00", 100002, "国机精工", "证券卖出", -100, "S2", 9],
        ])

        page = self.client.get("/analysis/reviews").get_data(as_text=True)
        self.assertIn("通富微电", page)
        self.assertNotIn("国机精工", page)
        self.assertIn("交易复盘", page)

        page = self.client.get("/analysis/reviews?start_date=2026-01-01&end_date=2026-12-31").get_data(as_text=True)
        self.assertIn("国机精工", page)
        self.assertIn("通富微电", page)

        page = self.client.get("/analysis/reviews?start_date=2020-01-01&end_date=2020-12-31").get_data(as_text=True)
        self.assertNotIn("通富微电", page)
        self.assertNotIn("国机精工", page)

        start30 = (today - timedelta(days=29)).strftime("%Y-%m-%d")
        page = self.client.get(f"/analysis/reviews?start_date={start30}&end_date={today.strftime('%Y-%m-%d')}").get_data(as_text=True)
        self.assertIn("通富微电", page)
        self.assertNotIn("国机精工", page)

        page = self.client.get("/analysis/reviews?status=all&start_date=2026-01-01&end_date=2026-12-31").get_data(as_text=True)
        self.assertIn("通富微电", page)
        self.assertIn("国机精工", page)

    def test_excursion_metrics_calculation(self):
        episode_id = self._import_workbook_rows([
            [20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
            [20260805, "10:00:00", 2156, "通富微电", "证券卖出", -100, "S1", 14],
        ])
        with self.module.db_connect() as db:
            user_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            for d, h, l in [("2026-08-01", 15, 9), ("2026-08-02", 16, 10), ("2026-08-03", 14, 8), ("2026-08-04", 13, 9), ("2026-08-05", 14, 10)]:
                db.execute(
                    "INSERT INTO daily_prices (stock_code, trade_date, open, high, low, close, volume, source, fetched_at) VALUES ('002156', ?, ?, ?, ?, ?, 0, 'test', '2026-08-21T00:00:00')",
                    (d, l, h, l, l),
                )
        with self.module.db_connect() as db:
            episode = [r for r in self.module.trade_episode_rows(db, user_id) if r["id"] == episode_id][0]
            metrics = self.module.calculate_excursion_metrics(db, user_id, episode)
        self.assertAlmostEqual(metrics["mfe"], 0.6, places=4)
        self.assertAlmostEqual(metrics["mae"], -0.2, places=4)
        self.assertAlmostEqual(metrics["capture_rate"], 400 / 600, places=4)
        self.assertAlmostEqual(metrics["highest_price"], 16)
        self.assertAlmostEqual(metrics["lowest_price"], 8)

        page = self.client.get(f"/analysis/trades/{episode_id}").get_data(as_text=True)
        self.assertIn("MFE 最大浮盈", page)
        self.assertIn("60.0%", page)
        self.assertIn("-20.0%", page)
        self.assertIn("利润捕获率", page)

    def test_excursion_capture_rate_edge_cases(self):
        lose = self._import_workbook_rows([
            [20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
            [20260803, "10:00:00", 2156, "通富微电", "证券卖出", -100, "S1", 9],
        ])
        with self.module.db_connect() as db:
            user_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            db.execute("INSERT INTO daily_prices (stock_code, trade_date, open, high, low, close, volume, source, fetched_at) VALUES ('002156', '2026-08-01', 10, 12, 9, 11, 0, 'test', 'x')")
            db.execute("INSERT INTO daily_prices (stock_code, trade_date, open, high, low, close, volume, source, fetched_at) VALUES ('002156', '2026-08-02', 11, 11, 8, 9, 0, 'test', 'x')")
            db.execute("INSERT INTO daily_prices (stock_code, trade_date, open, high, low, close, volume, source, fetched_at) VALUES ('002156', '2026-08-03', 9, 10, 8, 9, 0, 'test', 'x')")
        with self.module.db_connect() as db:
            episode = [r for r in self.module.trade_episode_rows(db, user_id) if r["id"] == lose][0]
            metrics = self.module.calculate_excursion_metrics(db, user_id, episode)
        self.assertEqual(metrics["capture_rate"], 0.0)

        flat = self._import_workbook_rows([
            [20260801, "09:30:00", 100002, "国机精工", "证券买入", 100, "B1", 10],
            [20260803, "10:00:00", 100002, "国机精工", "证券卖出", -100, "S1", 8],
        ])
        with self.module.db_connect() as db:
            db.execute("INSERT INTO daily_prices (stock_code, trade_date, open, high, low, close, volume, source, fetched_at) VALUES ('100002', '2026-08-01', 10, 9, 8, 9, 0, 'test', 'x')")
            db.execute("INSERT INTO daily_prices (stock_code, trade_date, open, high, low, close, volume, source, fetched_at) VALUES ('100002', '2026-08-03', 9, 9, 7, 8, 0, 'test', 'x')")
        with self.module.db_connect() as db:
            episode = [r for r in self.module.trade_episode_rows(db, user_id) if r["id"] == flat][0]
            metrics = self.module.calculate_excursion_metrics(db, user_id, episode)
        self.assertIsNone(metrics["capture_rate"])

        empty_page = self.client.get(f"/analysis/trades/{lose}").get_data(as_text=True)
        self.assertIn("同步行情并计算", empty_page)

    def test_trade_diagnosis_rules(self):
        episode = {"profit": 100.0, "status": "CLOSED", "sell_price": 14.0, "buy_price": 10.0, "stock_code": "002156", "sell_date": "2026-08-05", "id": 1}
        with self.module.db_connect() as db:
            codes = [item["code"] for item in self.module.trade_diagnosis(db, None, None, episode)]
        self.assertIn("NO_PRICE", codes)

        early = {"capture_rate": 0.3, "highest_price": 16.0, "lowest_price": 8.0}
        with self.module.db_connect() as db:
            codes = [item["code"] for item in self.module.trade_diagnosis(db, None, early, episode)]
        self.assertIn("EARLY_EXIT", codes)

        good = {"capture_rate": 0.9, "highest_price": 14.2, "lowest_price": 8.0}
        with self.module.db_connect() as db:
            codes = [item["code"] for item in self.module.trade_diagnosis(db, None, good, episode)]
        self.assertIn("GOOD_EXIT", codes)

        loss_episode = {"profit": -100.0, "status": "CLOSED", "sell_price": 14.0, "buy_price": 10.0, "stock_code": "002156", "sell_date": "2026-08-05", "id": 2}
        stop_review = {"stop_loss_price": 15.0, "expected_target_price": None}
        with self.module.db_connect() as db:
            codes = [item["code"] for item in self.module.trade_diagnosis(db, stop_review, early, loss_episode)]
        self.assertIn("STOP_VIOLATION", codes)

        target_review = {"stop_loss_price": None, "expected_target_price": 13.0}
        with self.module.db_connect() as db:
            codes = [item["code"] for item in self.module.trade_diagnosis(db, target_review, early, episode)]
        self.assertIn("TARGET_MET", codes)

        with self.module.db_connect() as db:
            codes = [item["code"] for item in self.module.trade_diagnosis(db, None, early, episode)]
        self.assertIn("NO_PLAN", codes)

        with self.module.db_connect() as db:
            db.execute("INSERT INTO daily_prices (stock_code, trade_date, open, high, low, close, volume, source, fetched_at) VALUES ('002156', '2026-08-10', 14, 15, 13, 15, 0, 'test', 'x')")
            codes = [item["code"] for item in self.module.trade_diagnosis(db, None, early, loss_episode)]
        self.assertIn("MISSED_REBOUND", codes)

        full_review = {"stop_loss_price": 12.0, "expected_target_price": 13.0}
        with self.module.db_connect() as db:
            codes = [item["code"] for item in self.module.trade_diagnosis(db, full_review, {"capture_rate": 0.8, "highest_price": 14.2, "lowest_price": 11.0}, episode)]
        self.assertNotIn("NO_PLAN", codes)
        self.assertIn("TARGET_MET", codes)

    def test_trading_system_stats(self):
        with self.module.db_connect() as db:
            user_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            stats = self.module.trading_system_stats(db, user_id)
            self.assertEqual(stats["total"], 0)
            self.assertIsNone(stats["avg_win"])
            self.assertIsNone(stats["profit_factor"])
            self.assertEqual(stats["max_win_streak"], 0)

        trades = [
            (2156, "通富微电", 10, 12, 20260802),
            (100002, "国机精工", 10, 9, 20260803),
            (300540, "蜀道装备", 10, 13, 20260804),
            (301697, "贝特利", 10, 14, 20260805),
            (600309, "万华化学", 10, 8, 20260806),
        ]
        for code, name, buy, sell, sell_day in trades:
            self._import_workbook_rows([
                [20260801, "09:30:00", code, name, "证券买入", 100, "B", buy],
                [sell_day, "10:00:00", code, name, "证券卖出", -100, "S", sell],
            ])
        with self.module.db_connect() as db:
            stats = self.module.trading_system_stats(db, user_id)
        self.assertEqual(stats["total"], 5)
        self.assertAlmostEqual(stats["avg_win"], 300.0, places=4)
        self.assertAlmostEqual(stats["avg_loss"], -150.0, places=4)
        self.assertAlmostEqual(stats["profit_factor"], 2.0, places=4)
        self.assertEqual(stats["max_win_streak"], 2)
        self.assertEqual(stats["max_loss_streak"], 1)

        summary = self.client.get("/analysis/summary").get_data(as_text=True)
        self.assertIn("我的交易系统", summary)
        self.assertIn("盈亏比", summary)
        self.assertIn("最大连续盈利", summary)
        self.assertIn("最大连续亏损", summary)

    def test_strategy_type_stats(self):
        with self.module.db_connect() as db:
            user_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            self.assertEqual(self.module.strategy_type_stats(db, user_id), [])

        win = self._import_workbook_rows([
            [20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
            [20260802, "10:00:00", 2156, "通富微电", "证券卖出", -100, "S1", 13],
        ])
        lose = self._import_workbook_rows([
            [20260803, "09:30:00", 100002, "国机精工", "证券买入", 100, "B2", 10],
            [20260804, "10:00:00", 100002, "国机精工", "证券卖出", -100, "S2", 9],
        ])
        self.client.post(f"/analysis/trades/{win}/review", data={"action": "complete", "trade_reason": "x", "judgement_result": "CORRECT", "reason_type": ["PULLBACK"], "sort": "recent"})
        self.client.post(f"/analysis/trades/{lose}/review", data={"action": "complete", "trade_reason": "x", "judgement_result": "WRONG", "reason_type": ["PULLBACK", "TECHNICAL_PATTERN"], "sort": "recent"})

        with self.module.db_connect() as db:
            stats = self.module.strategy_type_stats(db, user_id)
        by_type = {row["reason_type"]: row for row in stats}
        self.assertEqual(by_type["PULLBACK"]["trades"], 2)
        self.assertEqual(by_type["PULLBACK"]["wins"], 1)
        self.assertAlmostEqual(by_type["PULLBACK"]["total_profit"], 200.0)
        self.assertAlmostEqual(by_type["PULLBACK"]["win_rate"], 0.5)
        self.assertEqual(by_type["TECHNICAL_PATTERN"]["trades"], 1)
        self.assertAlmostEqual(by_type["TECHNICAL_PATTERN"]["total_profit"], -100.0)

        page = self.client.get("/analysis/strategy").get_data(as_text=True)
        self.assertIn("我的策略类型统计", page)
        self.assertIn("回调买入", page)
        self.assertIn("技术形态", page)
        self.assertIn("累计盈亏", page)

    def test_mistake_stats(self):
        with self.module.db_connect() as db:
            user_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            self.assertEqual(self.module.mistake_stats(db, user_id)["items"], [])

        ep1 = self._import_workbook_rows([
            [20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
            [20260802, "10:00:00", 2156, "通富微电", "证券卖出", -100, "S1", 13],
        ])
        ep2 = self._import_workbook_rows([
            [20260803, "09:30:00", 100002, "国机精工", "证券买入", 100, "B2", 10],
            [20260804, "10:00:00", 100002, "国机精工", "证券卖出", -100, "S2", 9],
        ])
        ep3 = self._import_workbook_rows([
            [20260805, "09:30:00", 300540, "蜀道装备", "证券买入", 100, "B3", 10],
            [20260806, "10:00:00", 300540, "蜀道装备", "证券卖出", -100, "S3", 8],
        ])
        self.client.post(f"/analysis/trades/{ep1}/review", data={"action": "complete", "trade_reason": "x", "judgement_result": "PARTIAL", "main_problem": "CHASE", "sort": "recent"})
        self.client.post(f"/analysis/trades/{ep2}/review", data={"action": "complete", "trade_reason": "x", "judgement_result": "WRONG", "main_problem": "CHASE", "sort": "recent"})
        self.client.post(f"/analysis/trades/{ep3}/review", data={"action": "complete", "trade_reason": "x", "judgement_result": "WRONG", "main_problem": "EXIT_TIMING", "sort": "recent"})

        with self.module.db_connect() as db:
            result = self.module.mistake_stats(db, user_id)
        self.assertEqual(result["total_reviewed"], 3)
        by_type = {row["main_problem"]: row for row in result["items"]}
        self.assertEqual(by_type["CHASE"]["count"], 2)
        self.assertAlmostEqual(by_type["CHASE"]["total_profit"], 200.0)
        self.assertAlmostEqual(by_type["CHASE"]["rate"], 2 / 3, places=4)
        self.assertEqual(by_type["EXIT_TIMING"]["count"], 1)
        self.assertAlmostEqual(by_type["EXIT_TIMING"]["total_profit"], -200.0)
        self.assertEqual(result["items"][0]["main_problem"], "CHASE")

        page = self.client.get("/analysis/discipline").get_data(as_text=True)
        self.assertIn("我的交易错误排行榜", page)
        self.assertIn("追涨", page)
        self.assertIn("卖点错误", page)

    def test_trade_episode_detail_is_user_isolated(self):
        self.client.post("/admin/preview", data={"statement": (self.make_statement(), "first.xlsx")}, content_type="multipart/form-data")
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW'").fetchone()
        self.client.post(f"/admin/import/{job['id']}")
        with self.module.db_connect() as db:
            episode_id = db.execute("SELECT id FROM trade_episodes").fetchone()["id"]
        self.client.post("/users/create", data={"name": "second", "next": "/analysis/stocks"})
        self.assertEqual(self.client.get(f"/analysis/trades/{episode_id}").status_code, 404)

    def test_paste_manual_executions(self):
        trades = "\n".join([
            "成交时间\t证券代码\t证券名称\t操作\t成交数量\t成交均价\t成交金额\t合同编号\t成交编号\t委托时间",
            "14:49:00\t002156\t通富微电\t证券卖出\t100\t63.810\t6381.000\t23921\t0105000070904721\t14:48:59",
            "09:35:29\t300540\t蜀道装备\t证券买入\t200\t29.890\t5978.000\t4322\t0101000007141777\t09:35:29",
        ])
        response = self.client.post(
            "/admin/manual/paste",
            data={"trade_date": "2026-08-21", "trades": trades},
            follow_redirects=True,
        )
        self.assertIn("新增 2 条", response.get_data(as_text=True))
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM executions WHERE source = 'BROKER_TODAY'").fetchone()[0], 2)
            sell = db.execute("SELECT * FROM executions WHERE stock_code = '002156'").fetchone()
            self.assertEqual(sell["action"], "SELL")
            self.assertEqual(sell["deal_id"], "0105000070904721")
            buy = db.execute("SELECT * FROM executions WHERE stock_code = '300540'").fetchone()
            self.assertAlmostEqual(buy["deal_price"], 29.89)

        response = self.client.post(
            "/admin/manual/paste",
            data={"trade_date": "2026-08-21", "trades": trades},
            follow_redirects=True,
        )
        self.assertIn("跳过重复 2 条", response.get_data(as_text=True))

    def test_statements_pagination_and_order(self):
        self.import_statement(days=25)
        page1 = self.client.get("/admin/statements")
        self.assertEqual(page1.status_code, 200)
        page1_text = page1.get_data(as_text=True)
        self.assertIn("第 1 / 2 页 · 共 25 条", page1_text)
        with self.module.db_connect() as db:
            first = db.execute("SELECT trade_date FROM executions ORDER BY trade_date DESC, trade_time DESC, id DESC LIMIT 1").fetchone()
        self.assertIn(first["trade_date"], page1_text)
        page2 = self.client.get("/admin/statements?page=2")
        self.assertEqual(page2.status_code, 200)
        page2_text = page2.get_data(as_text=True)
        self.assertIn("第 2 / 2 页 · 共 25 条", page2_text)
        with self.module.db_connect() as db:
            last = db.execute("SELECT trade_date FROM executions ORDER BY trade_date DESC, trade_time DESC, id DESC LIMIT 1 OFFSET 24").fetchone()
        self.assertIn(last["trade_date"], page2_text)

    def test_label_placement_no_overlap(self):
        labels = []
        for index in range(8):
            self.module.place_label(labels, 120.0 + index * 6, 100.0, f"+{index + 1},000.00", "#000")
        ys = [label[1] for label in labels]
        self.assertEqual(len(set(ys)), len(ys), "密集标注应上下错开不重叠")

    def test_stock_detail_single_sell_profit(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["成交日期", "成交时间", "证券代码", "证券名称", "操作", "成交数量", "成交编号", "成交价格", "成交金额", "佣金", "印花税", "其他杂费", "其他费", "过户费", "交易市场"])
        sheet.append([20260801, "09:31:00", 2156, "通富微电", "证券买入", 100, "B1", 10, 1000, 1, 0, 0, 0, 0, "深圳A股"])
        sheet.append([20260803, "13:00:00", 2156, "通富微电", "证券卖出", -50, "S1", 14, 700, 1, 2, 0, 0, 0, "深圳A股"])
        sheet.append([20260803, "14:00:00", 2156, "通富微电", "证券卖出", -50, "S2", 11, 550, 1, 2, 0, 0, 0, "深圳A股"])
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        self.client.post("/admin/preview", data={"statement": (output, "statement.xlsx")}, content_type="multipart/form-data")
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW' ORDER BY created_at DESC, rowid DESC LIMIT 1").fetchone()
        self.client.post(f"/admin/import/{job['id']}")

        detail = self.client.get("/analysis/stocks?code=002156")
        self.assertEqual(detail.status_code, 200)
        text = detail.get_data(as_text=True)
        self.assertIn("总交易笔数", text)
        self.assertIn("+243.00", text)
        self.assertIn("24.28%", text)
        self.assertNotIn("chart-wrap", text)
        self.assertNotIn("<svg", text)
        self.assertNotIn("查看复盘", text)
        self.assertNotIn("查看 FIFO", text)
        self.assertIn("stock-kline-trigger", text)
        kline = self.client.get("/analysis/stocks/kline/002156")
        self.assertEqual(kline.status_code, 200)
        self.assertEqual(kline.get_json()["stock_code"], "002156")
        self.assertEqual([point["action"] for point in kline.get_json()["bs_points"]], ["BUY", "SELL", "SELL"])

        with self.module.db_connect() as db:
            db.execute("DELETE FROM daily_prices WHERE stock_code = '002156'")
        with patch.object(self.module, "sync_history_codes", return_value=(0, ["002156：测试同步失败"])) as sync:
            empty_kline = self.client.get("/analysis/stocks/kline/002156")
        sync.assert_called_once_with(["002156"], full=True)
        self.assertTrue(empty_kline.get_json()["sync_attempted"])
        self.assertIn("测试同步失败", empty_kline.get_json()["sync_error"])

    def test_submenu_and_import_button(self):
        page = self.client.get("/admin/statements")
        text = page.get_data(as_text=True)
        self.assertIn('id="submenu-admin"', text)
        self.assertIn("交割单", text)
        self.assertIn("导入交割单", text)
        self.assertIn('class="database-backup"', text)
        self.assertIn("下载包含本机全部用户及所有业务数据的完整 SQLite 数据库", text)
        self.assertIn('href="/admin/database/export"', text)
        self.assertIn("重新计算 FIFO", text)
        self.assertIn("共 0 条成交记录", text)

        import_page = self.client.get("/admin/import").get_data(as_text=True)
        self.assertIn("选择文件、解析预览、确认写入", import_page)
        self.assertIn('id="file-title"', import_page)
        self.assertIn('id="preview-button"', import_page)
        self.assertIn("正在解析...", import_page)

        account_page = self.client.get("/admin/account").get_data(as_text=True)
        self.assertIn("账户资金", account_page)
        self.assertIn("记录当前资金", account_page)
        self.assertIn('class="database-backup"', account_page)
        self.assertIn('href="/admin/database/export"', account_page)

    def test_export_database_downloads_complete_valid_snapshot(self):
        with self.module.db_connect() as db:
            other_user_id = db.execute(
                "INSERT INTO users (name, created_at) VALUES (?, ?)",
                ("备份测试用户", datetime.now().isoformat(timespec="seconds")),
            ).lastrowid
            db.execute(
                "INSERT INTO account_snapshots (user_id, snapshot_date, total_assets, available_cash, note, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (other_user_id, "2026-08-23", 123456.78, 23456.78, "完整备份验证", "2026-08-23T12:00:00", "2026-08-23T12:00:00"),
            )

        response = self.client.get("/admin/database/export")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.mimetype, "application/vnd.sqlite3")
        self.assertEqual(response.headers["Cache-Control"], "no-store")
        self.assertRegex(
            response.headers["Content-Disposition"],
            r'attachment; filename=stocknotes-backup-\d{8}-\d{6}\.db',
        )
        self.assertTrue(response.data.startswith(b"SQLite format 3\x00"))

        with tempfile.TemporaryDirectory() as directory:
            backup_path = os.path.join(directory, "downloaded.db")
            with open(backup_path, "wb") as backup_file:
                backup_file.write(response.data)
            backup = sqlite3.connect(backup_path)
            try:
                self.assertEqual(backup.execute("PRAGMA user_version").fetchone()[0], self.module.SCHEMA_VERSION)
                self.assertEqual(backup.execute("PRAGMA foreign_key_check").fetchall(), [])
                users = {row[0] for row in backup.execute("SELECT name FROM users")}
                self.assertEqual(users, {"yutaoGS", "备份测试用户"})
                snapshot = backup.execute(
                    "SELECT total_assets, note FROM account_snapshots WHERE user_id = ?",
                    (other_user_id,),
                ).fetchone()
                self.assertEqual(snapshot, (123456.78, "完整备份验证"))
            finally:
                backup.close()

    def test_split_menus(self):
        root = self.client.get("/")
        self.assertEqual(root.status_code, 302)
        self.assertIn("/analysis/portfolio", root.headers["Location"])

        stocks = self.client.get("/analysis/stocks")
        self.assertEqual(stocks.status_code, 200)
        stocks_text = stocks.get_data(as_text=True)
        self.assertIn("个股分析", stocks_text)
        self.assertIn("还没有已平仓交易", stocks_text)

        summary = self.client.get("/analysis/summary")
        self.assertEqual(summary.status_code, 200)
        summary_text = summary.get_data(as_text=True)
        self.assertIn("汇总", summary_text)
        self.assertIn("月度盈亏", summary_text)
        self.assertIn("交易复盘", summary_text)

        portfolio = self.client.get("/analysis/portfolio")
        self.assertEqual(portfolio.status_code, 200)
        portfolio_text = portfolio.get_data(as_text=True)
        self.assertIn("持仓分析", portfolio_text)
        self.assertIn("不只看“持有什么”", portfolio_text)
        self.assertIn("当前没有持仓", portfolio_text)
        self.assertNotIn("仓位集中度", portfolio_text)
        self.assertNotIn("持仓效率", portfolio_text)

        watchlist = self.client.get("/analysis/watchlist")
        self.assertEqual(watchlist.status_code, 200)
        watchlist_text = watchlist.get_data(as_text=True)
        self.assertIn("重点观察", watchlist_text)
        self.assertIn('action="/analysis/watchlist/sync-prices"', watchlist_text)
        self.assertIn('action="/analysis/watchlist/sync-history"', watchlist_text)
        self.assertIn('id="watchlist-add-trigger"', watchlist_text)
        self.assertIn('id="watchlist-add-backdrop"', watchlist_text)
        self.assertNotIn('class="panel watchlist-add"', watchlist_text)

        reviews = self.client.get("/analysis/reviews")
        self.assertEqual(reviews.status_code, 200)
        reviews_text = reviews.get_data(as_text=True)
        self.assertIn("交易复盘", reviews_text)
        self.assertIn("时间范围", reviews_text)
        self.assertIn("7天", reviews_text)
        self.assertIn("15天", reviews_text)
        self.assertIn("30天", reviews_text)
        self.assertIn("查询", reviews_text)

        strategy = self.client.get("/analysis/strategy")
        self.assertEqual(strategy.status_code, 200)
        strategy_text = strategy.get_data(as_text=True)
        self.assertIn("策略验证", strategy_text)
        self.assertIn("没有符合当前条件的历史交易", strategy_text)

    def test_portfolio_analysis_uses_positions_and_latest_close(self):
        episode_id = self._import_workbook_rows([
            [20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
        ])

        page = self.client.get("/analysis/portfolio").get_data(as_text=True)
        self.assertIn("投入资金", page)
        self.assertIn("1,000.00", page)
        self.assertIn("暂无行情数据", page)
        self.assertIn('/analysis/portfolio/kline/002156', page)
        self.assertIn('id="portfolio-kline-backdrop"', page)
        self.assertIn("开盘：", page)
        self.assertIn("收盘：", page)
        self.assertIn("涨幅：", page)
        self.assertIn("MA5：", page)
        self.assertIn("MA10：", page)
        self.assertIn("MA20：", page)
        self.assertNotIn("持仓成本口径", page)

        stock_page = self.client.get("/analysis/stocks?code=002156").get_data(as_text=True)
        self.assertIn("当前持仓", stock_page)
        self.assertIn("查看成交与交易计划", stock_page)
        self.assertIn(f'/analysis/trades/{episode_id}', stock_page)

        with self.module.db_connect() as db:
            db.execute(
                """INSERT INTO daily_prices
                (stock_code, trade_date, open, high, low, close, volume, source, fetched_at)
                VALUES ('002156', '2026-08-20', 10, 10.5, 9.8, 10, 10000, 'test', '2026-08-20T15:01:00')"""
            )
            db.execute(
                """INSERT INTO daily_prices
                (stock_code, trade_date, open, high, low, close, volume, source, fetched_at)
                VALUES ('002156', '2026-08-21', 11, 12.5, 10.8, 12, 10000, 'test', '2026-08-21T15:01:00')"""
            )

        page = self.client.get("/analysis/portfolio").get_data(as_text=True)
        self.assertIn("总市值", page)
        self.assertIn("1,200.00", page)
        self.assertIn("+200.00", page)
        self.assertIn("20.0%", page)
        self.assertIn("涨跌幅", page)
        self.assertIn("+20.00%", page)
        self.assertIn("2026-08-21 · 15:01:00", page)
        self.assertNotIn("市值口径", page)
        self.assertIn("行情日期 2026-08-21", page)

        quotes = self.client.get("/analysis/portfolio/quotes").get_json()
        self.assertAlmostEqual(quotes["positions"][0]["change_rate"], 0.2)
        self.assertEqual(quotes["summary"]["price_fetched_at"], "2026-08-21T15:01:00")
        self.assertEqual(self.client.get("/analysis/portfolio/quotes").headers["Cache-Control"], "no-store")

        kline_response = self.client.get("/analysis/portfolio/kline/002156")
        self.assertEqual(kline_response.status_code, 200)
        kline = kline_response.get_json()
        self.assertEqual(kline["stock_code"], "002156")
        self.assertEqual(kline["stock_name"], "通富微电")
        self.assertEqual(kline["avg_cost"], 10)
        self.assertEqual(kline["data"][-1]["date"], "2026-08-21")
        self.assertEqual(kline["data"][-1]["close"], 12)
        self.assertEqual([point["action"] for point in kline["bs_points"]], ["BUY"])
        self.assertEqual(kline["bs_points"][0]["price"], 10)

        response = self.client.post(
            "/admin/account",
            data={
                "snapshot_date": "2026-08-21", "total_assets": "20000",
                "available_cash": "8000", "note": "收盘快照",
            },
            follow_redirects=True,
        )
        self.assertIn("已保存 2026-08-21 的账户资金快照", response.get_data(as_text=True))
        page = self.client.get("/analysis/portfolio").get_data(as_text=True)
        self.assertNotIn("20,000.00", page)
        self.assertNotIn("股票总仓位", page)
        self.assertNotIn("非股票资产", page)

    def test_account_snapshot_upsert_validation_and_user_isolation(self):
        response = self.client.post(
            "/admin/account",
            data={"snapshot_date": "2026-08-21", "total_assets": "10000", "available_cash": "12000"},
            follow_redirects=True,
        )
        self.assertIn("可用资金不能大于账户总资产", response.get_data(as_text=True))

        self.client.post(
            "/admin/account",
            data={"snapshot_date": "2026-08-21", "total_assets": "10000", "available_cash": "2000"},
        )
        self.client.post(
            "/admin/account",
            data={"snapshot_date": "2026-08-21", "total_assets": "12000", "available_cash": "3000"},
        )
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM account_snapshots").fetchone()[0], 1)
            self.assertEqual(db.execute("SELECT total_assets FROM account_snapshots").fetchone()[0], 12000)

        self.client.post("/users/create", data={"name": "account-second", "next": "/admin/account"})
        second_page = self.client.get("/admin/account").get_data(as_text=True)
        self.assertIn("还没有账户资金快照", second_page)
        self.assertEqual(self.client.get("/analysis/portfolio/kline/002156").status_code, 404)

    def test_portfolio_sync_with_empty_positions(self):
        response = self.client.post("/analysis/portfolio/sync-prices", follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("当前没有持仓，无需同步行情", response.get_data(as_text=True))

        self._import_workbook_rows([
            [20260820, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
        ])
        page = self.client.get("/analysis/portfolio").get_data(as_text=True)
        self.assertIn("当前持仓", page)
        self.assertNotIn("持仓时间", page)

    def test_watchlist_add_update_change_and_user_isolation(self):
        with patch.object(self.module, "sync_history_codes", return_value=(167, [])) as sync:
            response = self.client.post("/analysis/watchlist", data={"stock_code": "2156", "stock_name": "通富微电", "priority": "3", "note": "等待突破"}, follow_redirects=True)
        sync.assert_called_once_with(["002156"], full=True)
        self.assertIn("已加入重点观察：通富微电", response.get_data(as_text=True))
        with self.module.db_connect() as db:
            db.execute("INSERT INTO daily_prices (stock_code, trade_date, open, high, low, close, volume, source, fetched_at) VALUES ('002156', '2026-08-20', 10, 10, 9, 10, 100, 'test', '2026-08-21T15:00:00')")
            db.execute("INSERT INTO daily_prices (stock_code, trade_date, open, high, low, close, volume, source, fetched_at) VALUES ('002156', '2026-08-21', 10, 11, 10, 11, 100, 'test', '2026-08-22T15:00:00')")
        page = self.client.get("/analysis/watchlist").get_data(as_text=True)
        self.assertIn("+10.0%", page)
        self.assertIn("等待突破", page)
        self.assertIn("2026-08-21", page)
        quotes = self.client.get("/analysis/watchlist/quotes").get_json()
        self.assertAlmostEqual(quotes["items"][0]["change_rate"], 0.1)
        self.assertEqual(quotes["items"][0]["price_fetched_at"], "2026-08-22T15:00:00")
        self.assertEqual(self.client.get("/analysis/watchlist/quotes").headers["Cache-Control"], "no-store")
        self.assertIn('href="/analysis/watchlist?q=&amp;priority=all&amp;sort=change&amp;direction=desc"', page)
        self.assertNotIn("涨幅排名", page)
        self.assertIn('href="/analysis/watchlist"', page)
        self.assertIn("重置", page)
        self.assertIn("/analysis/watchlist/kline/002156", page)
        kline = self.client.get("/analysis/watchlist/kline/002156")
        self.assertEqual(kline.status_code, 200)
        self.assertEqual(kline.get_json()["data"][-1]["close"], 11)
        self.assertEqual(kline.get_json()["bs_points"], [])
        response = self.client.post("/analysis/watchlist", data={"stock_code": "002156", "stock_name": "通富微电", "priority": "1", "note": "更新理由"}, follow_redirects=True)
        self.assertIn("已更新重点观察：通富微电", response.get_data(as_text=True))
        with self.module.db_connect() as db:
            item = db.execute("SELECT * FROM watchlist_stocks").fetchone()
            self.assertEqual(item["priority"], 1)
            self.assertEqual(item["note"], "更新理由")
        page = self.client.get("/analysis/watchlist").get_data(as_text=True)
        self.assertIn("watch-edit-trigger", page)
        self.assertIn('data-priority="1"', page)
        self.client.post("/users/create", data={"name": "watch-second", "next": "/analysis/watchlist"})
        self.assertEqual(self.client.get("/analysis/watchlist").get_data(as_text=True).count("通富微电"), 0)
        self.assertEqual(self.client.get("/analysis/watchlist/kline/002156").status_code, 404)

    def test_watchlist_sync_uses_incremental_date_range(self):
        with patch.object(self.module, "sync_history_codes", return_value=(0, [])):
            self.client.post(
                "/analysis/watchlist",
                data={"stock_code": "2156", "stock_name": "通富微电", "priority": "2"},
            )
        today = date.today()
        expected_initial = (today - timedelta(days=250)).isoformat()
        latest_date = today - timedelta(days=1)
        expected_incremental = (latest_date - timedelta(days=7)).isoformat()

        with patch.object(self.module, "fetch_daily_prices", return_value=(None, None)) as fetch:
            response = self.client.post("/analysis/watchlist/sync-history")
            self.assertEqual(response.status_code, 302)
            fetch.assert_called_once_with("002156", expected_initial, today.isoformat())

        with self.module.db_connect() as db:
            db.execute(
                """INSERT INTO daily_prices
                (stock_code, trade_date, close, source, fetched_at)
                VALUES (?, ?, 12, 'test', '2026-08-21T15:00:00')""",
                ("002156", latest_date.isoformat()),
            )

        with patch.object(self.module, "fetch_daily_prices", return_value=(None, None)) as fetch:
            response = self.client.post("/analysis/watchlist/sync-history")
            self.assertEqual(response.status_code, 302)
            fetch.assert_called_once_with("002156", expected_incremental, today.isoformat())

    def test_watchlist_sync_fetches_stocks_in_parallel(self):
        with patch.object(self.module, "sync_history_codes", return_value=(0, [])):
            for stock_code in ("000001", "000002", "000003", "000004"):
                self.client.post(
                    "/analysis/watchlist",
                    data={"stock_code": stock_code, "stock_name": stock_code, "priority": "2"},
                )

        lock = threading.Lock()
        active = 0
        max_active = 0

        def fetch(*args):
            nonlocal active, max_active
            with lock:
                active += 1
                max_active = max(max_active, active)
            time.sleep(0.05)
            with lock:
                active -= 1
            return None, None

        with patch.object(self.module, "fetch_daily_prices", side_effect=fetch) as mocked_fetch:
            response = self.client.post("/analysis/watchlist/sync-history")

        self.assertEqual(response.status_code, 302)
        self.assertEqual(mocked_fetch.call_count, 4)
        self.assertGreater(max_active, 1)

    def test_watchlist_sync_supports_ajax_response(self):
        with patch.object(self.module, "sync_history_codes", return_value=(0, [])):
            self.client.post(
                "/analysis/watchlist",
                data={"stock_code": "002156", "stock_name": "通富微电", "priority": "2"},
            )
        with patch.object(self.module, "sync_realtime_codes", return_value=(1, [])):
            response = self.client.post(
                "/analysis/watchlist/sync-prices",
                headers={"X-Requested-With": "XMLHttpRequest"},
            )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json(), {"message": "已同步 1 条实时行情。", "category": "success"})

    def test_watchlist_sync_prefers_batch_realtime_prices(self):
        quote = {
            "002156": {
                "trade_date": "2026-08-21", "open": 10.0, "high": 12.0, "low": 9.5,
                "close": 11.5, "volume": 1000.0, "amount": 11500.0,
                "fetched_at": "2026-08-21T15:00:00",
            }
        }
        with patch.object(self.module, "fetch_realtime_prices", return_value=quote) as realtime, \
             patch.object(self.module, "fetch_daily_prices") as history:
            synced, failed = self.module.sync_realtime_codes(["002156"])

        realtime.assert_called_once_with(["002156"])
        history.assert_not_called()
        self.assertEqual((synced, failed), (1, []))
        with self.module.db_connect() as db:
            price = db.execute("SELECT * FROM daily_prices WHERE stock_code = '002156'").fetchone()
        self.assertEqual(price["close"], 11.5)
        self.assertEqual(price["source"], "tencent-realtime")

    def _intraday_rebound_sample(self):
        prices = [
            ("13:50", 0.723, 1000), ("13:51", 0.709, 2000), ("13:52", 0.707, 3000),
            ("13:53", 0.705, 4000), ("13:54", 0.707, 5000), ("13:55", 0.709, 6000),
            ("13:56", 0.711, 7000), ("13:57", 0.709, 8000), ("13:58", 0.708, 9000),
            ("13:59", 0.707, 10000), ("14:00", 0.709, 11000), ("14:01", 0.710, 12000),
            ("14:02", 0.711, 13000), ("14:03", 0.710, 14000), ("14:04", 0.711, 15000),
            ("14:05", 0.713, 18000),
        ]
        return [
            {"quote_minute": f"2026-08-24T{minute}", "price": price, "volume": volume}
            for minute, price, volume in prices
        ]

    def test_intraday_rebound_requires_breakout_after_higher_low(self):
        samples = self._intraday_rebound_sample()
        before_breakout = self.module.evaluate_intraday_rebound(samples[:-1])
        result = self.module.evaluate_intraday_rebound(samples)

        self.assertFalse(before_breakout["matched"])
        self.assertTrue(result["matched"])
        self.assertEqual(result["trough_price"], 0.705)
        self.assertEqual(result["breakout_price"], 0.711)
        self.assertGreater(result["higher_low"], result["trough_price"])
        self.assertGreaterEqual(result["volume_multiple"], 1.5)

    def test_intraday_rebound_alert_persists_minutes_and_deduplicates(self):
        now = "2026-08-24T14:05:00"
        with self.module.db_connect() as db:
            user_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            db.execute(
                """INSERT INTO watchlist_stocks
                (user_id, stock_code, stock_name, priority, note, created_at, updated_at)
                VALUES (?, '159516', '半导体设备ETF', 3, '', ?, ?)""",
                (user_id, now, now),
            )
            for sample in self._intraday_rebound_sample():
                db.execute(
                    """INSERT INTO intraday_quotes
                    (stock_code, trade_date, quote_minute, price, volume, source, fetched_at)
                    VALUES ('159516', '2026-08-24', ?, ?, ?, 'test', ?)""",
                    (sample["quote_minute"], sample["price"], sample["volume"], sample["quote_minute"]),
                )
        quote = {
            "159516": {
                "trade_date": "2026-08-24", "quote_time": now, "open": 0.73, "high": 0.738,
                "low": 0.705, "close": 0.713, "previous_close": 0.729, "volume": 18000,
                "amount": 12834, "fetched_at": now,
            }
        }
        with self.module.db_connect() as db:
            self.module.evaluate_intraday_rebound_alerts(db, quote)
            self.module.evaluate_intraday_rebound_alerts(db, quote)
            notifications = db.execute(
                "SELECT title, content FROM notifications WHERE stock_code = '159516'"
            ).fetchall()
        self.assertEqual(len(notifications), 1)
        self.assertIn("日内反弹", notifications[0]["title"])
        self.assertIn("突破 0.711", notifications[0]["content"])

    def _exhaustion_sample(self):
        return [
            {"trade_date": "2026-08-17", "open": 133.50, "high": 135.79, "low": 129.50, "close": 135.50, "volume": 23473900},
            {"trade_date": "2026-08-18", "open": 135.75, "high": 136.55, "low": 128.21, "close": 131.70, "volume": 30619800},
            {"trade_date": "2026-08-19", "open": 126.00, "high": 127.15, "low": 118.53, "close": 118.53, "volume": 31212500},
            {"trade_date": "2026-08-20", "open": 118.53, "high": 121.84, "low": 115.17, "close": 118.28, "volume": 29231900},
        ]

    def test_three_day_dip_exhaustion_and_strong_reversal_samples(self):
        result = self.module.evaluate_three_day_dip(self._exhaustion_sample(), enforce_volume=True)
        self.assertTrue(result["matched"])
        self.assertEqual(result["pattern_type"], "EXHAUSTION")
        reversal = [
            {"trade_date": "2026-06-10", "open": 145.08, "high": 158.80, "low": 141.22, "close": 154.96, "volume": 34759500},
            {"trade_date": "2026-06-11", "open": 150.01, "high": 157.88, "low": 139.46, "close": 139.46, "volume": 29529300},
            {"trade_date": "2026-06-12", "open": 137.07, "high": 149.87, "low": 125.51, "close": 129.80, "volume": 54606600},
            {"trade_date": "2026-06-15", "open": 128.51, "high": 138.88, "low": 119.00, "close": 137.30, "volume": 37989800},
        ]
        result = self.module.evaluate_three_day_dip(reversal, enforce_volume=True)
        self.assertTrue(result["matched"])
        self.assertEqual(result["pattern_type"], "STRONG_REVERSAL")

    def test_three_day_dip_shadow_stop_is_candidate_only(self):
        prices = [
            {"trade_date": "2026-08-17", "open": 40.99, "high": 45.09, "low": 40.98, "close": 45.09, "volume": 428071},
            {"trade_date": "2026-08-18", "open": 45.54, "high": 45.99, "low": 42.65, "close": 43.67, "volume": 476475},
            {"trade_date": "2026-08-19", "open": 41.92, "high": 42.44, "low": 39.30, "close": 39.36, "volume": 455993},
            {"trade_date": "2026-08-20", "open": 39.74, "high": 41.13, "low": 38.60, "close": 39.25, "volume": 426727},
        ]
        result = self.module.evaluate_three_day_dip(prices, enforce_volume=True)
        self.assertTrue(result["matched"])
        self.assertEqual(result["pattern_type"], "SHADOW_STOP")
        self.assertGreater(result["lower_shadow_body_ratio"], 1)

    def test_three_day_dip_allows_signal_low_one_percent_above_prior_low(self):
        prices = [
            {"trade_date": "2026-08-18", "open": 72.01, "high": 75.50, "low": 71.11, "close": 72.94, "volume": 1184806},
            {"trade_date": "2026-08-19", "open": 68.80, "high": 70.60, "low": 65.68, "close": 66.16, "volume": 1048449},
            {"trade_date": "2026-08-20", "open": 68.00, "high": 68.24, "low": 63.58, "close": 64.45, "volume": 843824},
            {"trade_date": "2026-08-21", "open": 64.51, "high": 68.60, "low": 63.91, "close": 68.14, "volume": 835712},
        ]
        result = self.module.evaluate_three_day_dip(prices, enforce_volume=True)
        self.assertFalse(result["new_low_ok"])
        self.assertTrue(result["low_position_ok"])
        self.assertTrue(result["common_ok"])
        self.assertTrue(result["matched"])
        self.assertEqual(result["pattern_type"], "STRONG_REVERSAL")
        self.assertAlmostEqual(result["signal_low_above_prior_ratio"], 63.91 / 63.58 - 1)

        prices[-1]["low"] = 66.20
        result = self.module.evaluate_three_day_dip(prices, enforce_volume=True)
        self.assertFalse(result["low_position_ok"])
        self.assertFalse(result["common_ok"])

    def test_three_day_dip_strong_reversal_allows_four_percent_above_prior_low(self):
        prices = [
            {"trade_date": "2025-11-17", "open": 16.55, "high": 18.24, "low": 16.10, "close": 18.24, "volume": 181425400},
            {"trade_date": "2025-11-21", "open": 15.05, "high": 17.97, "low": 14.59, "close": 14.59, "volume": 227157400},
            {"trade_date": "2025-11-24", "open": 14.00, "high": 14.75, "low": 11.92, "close": 12.58, "volume": 210811600},
            {"trade_date": "2025-11-25", "open": 12.58, "high": 14.09, "low": 12.34, "close": 13.43, "volume": 192976400},
        ]
        result = self.module.evaluate_three_day_dip(prices, enforce_volume=True)
        self.assertTrue(result["low_position_ok"])
        self.assertTrue(result["matched"])
        self.assertEqual(result["pattern_type"], "STRONG_REVERSAL")
        self.assertAlmostEqual(result["signal_low_above_prior_ratio"], 12.34 / 11.92 - 1)

    def test_three_day_dip_realtime_alerts_deduplicate_candidate_and_confirmation(self):
        now = "2026-08-14T14:30:00"
        with self.module.db_connect() as db:
            user_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            db.execute(
                """INSERT INTO watchlist_stocks
                (user_id, stock_code, stock_name, priority, note, created_at, updated_at)
                VALUES (?, '600482', '中国动力', 3, '', ?, ?)""",
                (user_id, now, now),
            )
            for row in self._exhaustion_sample()[:-1]:
                db.execute(
                    """INSERT INTO daily_prices
                    (stock_code, trade_date, open, high, low, close, volume, source, fetched_at)
                    VALUES ('600482', ?, ?, ?, ?, ?, ?, 'test', ?)""",
                    (row["trade_date"], row["open"], row["high"], row["low"], row["close"], row["volume"], now),
                )

        candidate = {"600482": dict(self._exhaustion_sample()[-1], fetched_at=now, amount=1)}
        with patch.object(self.module, "fetch_realtime_prices", return_value=candidate):
            self.module.sync_realtime_codes(["600482"])
            self.module.sync_realtime_codes(["600482"])
        with self.module.db_connect() as db:
            stages = [row["stage"] for row in db.execute("SELECT stage FROM notifications ORDER BY id")]
        self.assertEqual(stages, ["CANDIDATE"])

        confirmed = {"600482": dict(candidate["600482"], fetched_at="2026-08-14T15:00:00")}
        with patch.object(self.module, "fetch_realtime_prices", return_value=confirmed):
            self.module.sync_realtime_codes(["600482"])
            self.module.sync_realtime_codes(["600482"])
        with self.module.db_connect() as db:
            stages = [row["stage"] for row in db.execute("SELECT stage FROM notifications ORDER BY id")]
        self.assertEqual(stages, ["CANDIDATE", "CONFIRMED"])

    def test_notification_api_is_user_scoped(self):
        now = "2026-08-14T15:00:00"
        with self.module.db_connect() as db:
            first_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            second_id = db.execute("INSERT INTO users (name, created_at) VALUES ('notify-second', ?)", (now,)).lastrowid
            alert_type_id = db.execute("SELECT id FROM alert_types WHERE code = ?", (self.module.THREE_DAY_DIP_CODE,)).fetchone()["id"]
            notification_id = db.execute(
                """INSERT INTO notifications
                (user_id, alert_type_id, stock_code, stock_name, stage, title, content,
                 details_json, quote_time, created_at, dedupe_key)
                VALUES (?, ?, '600482', '中国动力', 'CANDIDATE', '测试提醒', '只属于第二用户', '{}', ?, ?, 'scope-test')""",
                (second_id, alert_type_id, now, now),
            ).lastrowid

        response = self.client.get("/api/notifications")
        self.assertEqual(response.get_json()["unread_count"], 0)
        self.assertEqual(self.client.post(f"/api/notifications/{notification_id}/read").status_code, 404)
        with self.client.session_transaction() as session:
            session["user_id"] = second_id
        response = self.client.get("/api/notifications")
        self.assertEqual(response.get_json()["unread_count"], 1)
        self.assertEqual(response.get_json()["notifications"][0]["title"], "测试提醒")
        self.assertEqual(self.client.post(f"/api/notifications/{notification_id}/read").status_code, 200)
        self.assertEqual(self.client.get("/api/notifications").get_json()["unread_count"], 0)

    def _insert_notification(self, db, user_id, index, read=False):
        now = "2026-08-14T15:00:00"
        alert_type_id = db.execute("SELECT id FROM alert_types WHERE code = ?", (self.module.THREE_DAY_DIP_CODE,)).fetchone()["id"]
        return db.execute(
            """INSERT INTO notifications
            (user_id, alert_type_id, stock_code, stock_name, stage, title, content,
             details_json, quote_time, created_at, read_at, dedupe_key)
            VALUES (?, ?, '600482', '中国动力', ?, ?, ?, '{}', ?, ?, ?, ?)""",
            (
                user_id, alert_type_id, "CANDIDATE" if index % 2 else "CONFIRMED",
                f"通知标题 {index}", f"通知内容 {index}",
                now, now, now if read else None, f"page-key-{user_id}-{index}",
            ),
        ).lastrowid

    def test_notifications_page_pagination_and_read(self):
        now = "2026-08-14T15:00:00"
        with self.module.db_connect() as db:
            user_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            for index in range(25):
                self._insert_notification(db, user_id, index)

        page = self.client.get("/analysis/notifications").get_data(as_text=True)
        self.assertIn("共 25 条", page)
        self.assertIn("第 1 / 2 页", page)
        self.assertIn("下一页", page)
        self.assertIn("通知标题 24", page)
        self.assertIn("通知标题 5", page)
        second = self.client.get("/analysis/notifications?page=2").get_data(as_text=True)
        self.assertIn("通知标题 0", second)
        self.assertIn("通知标题 4", second)
        self.assertNotIn("通知标题 5", second)
        self.assertIn("第 2 / 2 页", second)

        with self.module.db_connect() as db:
            newest = db.execute("SELECT id FROM notifications ORDER BY created_at DESC, id DESC LIMIT 1").fetchone()["id"]
        self.assertEqual(self.client.post(f"/api/notifications/{newest}/read").status_code, 200)
        page = self.client.get("/analysis/notifications").get_data(as_text=True)
        self.assertIn("24 条未读", page)
        first_item = page.split('class="notification-page-item', 1)
        self.assertEqual(len(first_item), 2)
        self.assertNotIn("unread", first_item[1].split('"', 1)[0])

    def test_notifications_page_user_isolated_and_empty_state(self):
        now = "2026-08-14T15:00:00"
        with self.module.db_connect() as db:
            first_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            second_id = db.execute("INSERT INTO users (name, created_at) VALUES ('page-second', ?)", (now,)).lastrowid
            self._insert_notification(db, second_id, 1)

        page = self.client.get("/analysis/notifications").get_data(as_text=True)
        self.assertIn("还没有提醒通知", page)
        self.assertNotIn("通知标题 1", page)
        with self.client.session_transaction() as session:
            session["user_id"] = second_id
        page = self.client.get("/analysis/notifications").get_data(as_text=True)
        self.assertIn("通知标题 1", page)

    def test_alert_type_admin_updates_parameters_without_builtin_sample(self):
        page = self.client.get("/admin/alert-types/three-day-dip").get_data(as_text=True)
        self.assertIn("衰竭止跌型", page)
        self.assertIn("强修复型", page)
        self.assertNotIn("中国动力", page)

        response = self.client.post(
            "/admin/alert-types/three-day-dip",
            data={
                "enabled": "on", "decline_days": "2", "require_bearish_candles": "on",
                "require_declining_closes": "on", "require_signal_day_new_low": "on",
                "min_decline_percent": "9", "max_volume_percent": "95",
                "exhaustion_min_repair_percent": "2", "exhaustion_max_body_percent": "10",
                "exhaustion_min_change_percent": "-1.5", "exhaustion_max_change_percent": "3",
                "reversal_min_decline_percent": "11", "reversal_min_repair_percent": "7",
                "reversal_min_recovery_percent": "85", "reversal_min_change_percent": "5",
                "reversal_max_volume_percent": "75",
                "intraday_candidate_enabled": "on", "close_confirmation_enabled": "on",
            },
            follow_redirects=True,
        )
        text = response.get_data(as_text=True)
        self.assertIn("已保存三日低吸提醒参数", text)
        self.assertIn("9.0", text)

    def test_intraday_rebound_management_menu_and_parameters(self):
        page = self.client.get("/admin/alert-types/intraday-rebound").get_data(as_text=True)
        self.assertIn("日内反弹", page)
        self.assertIn("分钟级提醒", page)
        self.assertIn("急跌幅度至少", page)

        response = self.client.post(
            "/admin/alert-types/intraday-rebound",
            data={
                "enabled": "on", "lookback_minutes": "90", "min_drop_percent": "2.5",
                "min_rebound_percent": "1.2", "min_trough_age_minutes": "6", "min_volume_multiple": "2",
            },
            follow_redirects=True,
        )
        self.assertIn("已保存日内反弹提醒参数", response.get_data(as_text=True))
        with self.module.db_connect() as db:
            row = db.execute("SELECT params_json FROM alert_types WHERE code = ?", (self.module.INTRADAY_REBOUND_CODE,)).fetchone()
        params = json.loads(row["params_json"])
        self.assertEqual(params["lookback_minutes"], 90)
        self.assertEqual(params["min_drop_ratio"], 0.025)
        self.assertEqual(params["min_volume_multiple"], 2.0)

    def test_three_day_dip_backtest_finds_hits_without_writing_notifications(self):
        now = "2026-08-15T10:00:00"
        with self.module.db_connect() as db:
            user_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            stock = db.execute(
                """INSERT INTO watchlist_stocks
                (user_id, stock_code, stock_name, priority, note, created_at, updated_at)
                VALUES (?, '600482', '中国动力', 3, '', ?, ?) RETURNING stock_code, stock_name""",
                (user_id, now, now),
            ).fetchone()
            for row in self._exhaustion_sample():
                db.execute(
                    """INSERT INTO daily_prices
                    (stock_code, trade_date, open, high, low, close, volume, source, fetched_at)
                    VALUES ('600482', ?, ?, ?, ?, ?, ?, 'test', ?)""",
                    (row["trade_date"], row["open"], row["high"], row["low"], row["close"], row["volume"], now),
                )
            result = self.module.backtest_three_day_dip(
                db, stock, "2026-08-20", "2026-08-20", self.module.THREE_DAY_DIP_DEFAULT_PARAMS,
            )

        self.assertEqual(result["scanned_days"], 1)
        self.assertEqual([hit["signal_date"] for hit in result["hits"]], ["2026-08-20"])
        self.assertEqual(result["hits"][0]["pattern_type"], "EXHAUSTION")
        response = self.client.get(
            "/admin/alert-types/three-day-dip?run_test=1&stock_code=600482&start_date=2026-08-20&end_date=2026-08-20"
        )
        text = response.get_data(as_text=True)
        self.assertIn("重点观察历史命中测试", text)
        self.assertIn("衰竭止跌型", text)
        self.assertIn("衰竭止跌型", text)
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM notifications").fetchone()[0], 0)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM alert_rule_states").fetchone()[0], 0)

    def test_three_day_dip_backtest_handles_insufficient_data_and_user_scope(self):
        now = "2026-08-15T10:00:00"
        with self.module.db_connect() as db:
            first_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            second_id = db.execute("INSERT INTO users (name, created_at) VALUES ('backtest-second', ?)", (now,)).lastrowid
            db.execute(
                """INSERT INTO watchlist_stocks
                (user_id, stock_code, stock_name, priority, note, created_at, updated_at)
                VALUES (?, '000001', '平安银行', 2, '', ?, ?)""",
                (first_id, now, now),
            )
            db.execute(
                """INSERT INTO watchlist_stocks
                (user_id, stock_code, stock_name, priority, note, created_at, updated_at)
                VALUES (?, '600482', '中国动力', 2, '', ?, ?)""",
                (second_id, now, now),
            )
            db.execute(
                """INSERT INTO daily_prices
                (stock_code, trade_date, open, high, low, close, source, fetched_at)
                VALUES ('000001', '2026-08-14', 10, 11, 9, 10, 'test', ?)""",
                (now,),
            )

        insufficient = self.client.get(
            "/admin/alert-types/three-day-dip?run_test=1&stock_code=000001&start_date=2026-08-14&end_date=2026-08-14"
        ).get_data(as_text=True)
        self.assertIn("历史行情不足，无法完成测试", insufficient)
        scoped = self.client.get(
            "/admin/alert-types/three-day-dip?run_test=1&stock_code=600482&start_date=2026-08-14&end_date=2026-08-14"
        ).get_data(as_text=True)
        self.assertIn("请选择当前用户重点观察中的股票", scoped)
        self.assertNotIn("中国动力 · 600482", scoped)

    def test_three_day_dip_backtest_validates_date_range(self):
        now = "2026-08-15T10:00:00"
        with self.module.db_connect() as db:
            user_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            db.execute(
                """INSERT INTO watchlist_stocks
                (user_id, stock_code, stock_name, priority, note, created_at, updated_at)
                VALUES (?, '000001', '平安银行', 2, '', ?, ?)""",
                (user_id, now, now),
            )
        text = self.client.get(
            "/admin/alert-types/three-day-dip?run_test=1&stock_code=000001&start_date=2026-08-15&end_date=2026-08-14"
        ).get_data(as_text=True)
        self.assertIn("开始日期不能晚于截止日期", text)

    def test_three_day_dip_pool_manual_sample_review_and_outcomes(self):
        now = "2026-08-23T15:00:00"
        prices = self._exhaustion_sample() + [
            {"trade_date": "2026-08-21", "open": 118.83, "high": 129.98, "low": 117.06, "close": 128.06, "volume": 34277800},
        ]
        with self.module.db_connect() as db:
            for row in prices:
                db.execute(
                    """INSERT INTO daily_prices
                    (stock_code, trade_date, open, high, low, close, volume, source, fetched_at)
                    VALUES ('002851', ?, ?, ?, ?, ?, ?, 'test', ?)""",
                    (row["trade_date"], row["open"], row["high"], row["low"], row["close"], row["volume"], now),
                )
        response = self.client.post(
            "/admin/alert-types/three-day-dip/pool",
            data={"stock_code": "002851", "stock_name": "麦格米特", "signal_date": "2026-08-20", "note": "金标准"},
            follow_redirects=True,
        )
        text = response.get_data(as_text=True)
        self.assertIn("已加入三日低吸池", text)
        self.assertIn("麦格米特", text)
        self.assertIn("8.3%", text)
        with self.module.db_connect() as db:
            sample = db.execute("SELECT * FROM alert_signal_samples").fetchone()
            outcome = db.execute("SELECT * FROM alert_signal_outcomes").fetchone()
        self.assertEqual(sample["review_status"], "CONFIRMED")
        self.assertEqual(sample["source"], "MANUAL")
        self.assertAlmostEqual(outcome["day_1_close_return"], 128.06 / 118.28 - 1)

        self.assertEqual(self.client.post(
            f"/admin/alert-types/three-day-dip/pool/{sample['id']}/review",
            data={"review_status": "REJECTED", "note": "复核否决"},
        ).status_code, 302)
        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT review_status FROM alert_signal_samples").fetchone()[0], "REJECTED")

    def test_fetch_realtime_prices_parses_tencent_response(self):
        fields = [""] * 88
        fields[1] = "通富微电"
        fields[2] = "002156"
        fields[3] = "11.50"
        fields[4] = "11.00"
        fields[5] = "11.10"
        fields[30] = "20260821150000"
        fields[33] = "12.00"
        fields[34] = "10.80"
        fields[35] = "11.50/1000/11500"
        fields[36] = "1000"
        payload = ('v_sz002156="' + "~".join(fields) + '";').encode("gbk")

        class Response:
            def __enter__(self):
                return self

            def __exit__(self, *args):
                return None

            def read(self):
                return payload

        with patch.object(self.module, "urlopen", return_value=Response()):
            quotes = self.module.fetch_realtime_prices(["002156"])

        quote = quotes["002156"]
        self.assertEqual(quote["trade_date"], "2026-08-21")
        self.assertEqual(quote["open"], 11.1)
        self.assertEqual(quote["high"], 12.0)
        self.assertEqual(quote["low"], 10.8)
        self.assertEqual(quote["close"], 11.5)
        self.assertEqual(quote["volume"], 1000.0)
        self.assertEqual(quote["amount"], 11500.0)
        self.assertEqual(quote["previous_close"], 11.0)

    def test_kline_volume_converts_realtime_shares_to_lots(self):
        with self.module.db_connect() as db:
            db.execute(
                """INSERT INTO daily_prices
                (stock_code, trade_date, open, high, low, close, volume, source, fetched_at)
                VALUES ('000657', '2026-08-24', 69, 70, 64, 66.76, 896396, 'tencent-realtime', '2026-08-24T15:03:06')"""
            )
            row = db.execute("SELECT volume, source FROM daily_prices WHERE stock_code = '000657'").fetchone()
        self.assertEqual(self.module.kline_volume_lots(row), 8963.96)

    def test_fetch_market_indexes_parses_index_quotes(self):
        fields = [""] * 88
        fields[2] = "sh000001"
        fields[3] = "3500.00"
        fields[4] = "3490.00"
        fields[5] = "3480.00"
        fields[30] = "20260821150000"
        fields[33] = "3510.00"
        fields[34] = "3470.00"
        fields[35] = "100000/1000/100000000"
        fields[36] = "1000000"
        payload = ('v_sh000001="' + "~".join(fields) + '";').encode("gbk")

        class Response:
            def __enter__(self):
                return self

            def __exit__(self, *args):
                return None

            def read(self):
                return payload

        with patch.object(self.module, "urlopen", return_value=Response()):
            indexes = self.module.fetch_market_indexes()

        self.assertEqual(len(indexes), 1)
        self.assertEqual(indexes[0]["name"], "上证指数")
        self.assertEqual(indexes[0]["price"], 3500)
        self.assertEqual(indexes[0]["change"], 10)
        self.assertAlmostEqual(indexes[0]["change_rate"], 10 / 3490)

    def test_fetch_market_indexes_parses_multiple_records_with_newlines(self):
        def record(code, close, previous):
            fields = [""] * 88
            fields[2] = code[2:]
            fields[3] = str(close)
            fields[4] = str(previous)
            fields[5] = str(previous)
            fields[30] = "20260821150000"
            fields[33] = str(close)
            fields[34] = str(previous)
            fields[35] = "100/1000/100000"
            fields[36] = "1000"
            return "v_" + code + '=\"' + "~".join(fields) + '\"'

        payload = (record("sh000001", 3500, 3490) + ";\n" + record("sz399001", 14000, 13900) + ";\n" + record("sz399006", 3500, 3400) + ";\n" + record("sh000688", 1600, 1590) + ";").encode("gbk")

        class Response:
            def __enter__(self):
                return self

            def __exit__(self, *args):
                return None

            def read(self):
                return payload

        with patch.object(self.module, "urlopen", return_value=Response()):
            indexes = self.module.fetch_market_indexes()

        self.assertEqual([item["code"] for item in indexes], list(self.module.MARKET_INDEXES)[:4])

    def test_portfolio_indexes_api_returns_error_without_name_error(self):
        with patch.object(self.module, "fetch_market_indexes", return_value=[]):
            response = self.client.get("/analysis/portfolio/indexes")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["data"], [])

    def test_auto_sync_trading_windows(self):
        self.assertEqual(self.module.AUTO_REALTIME_SYNC_INTERVAL_SECONDS, 30)
        self.assertEqual(self.module.HOT_SECTORS_CACHE_TTL_SECONDS, 30)
        self.assertFalse(self.module.is_auto_sync_time(datetime(2026, 8, 24, 9, 14)))
        self.assertTrue(self.module.is_auto_sync_time(datetime(2026, 8, 24, 9, 15)))
        self.assertTrue(self.module.is_auto_sync_time(datetime(2026, 8, 24, 11, 30, 59)))
        self.assertFalse(self.module.is_auto_sync_time(datetime(2026, 8, 24, 11, 31)))
        self.assertFalse(self.module.is_auto_sync_time(datetime(2026, 8, 24, 12, 59)))
        self.assertTrue(self.module.is_auto_sync_time(datetime(2026, 8, 24, 13, 0)))
        self.assertTrue(self.module.is_auto_sync_time(datetime(2026, 8, 24, 15, 0, 59)))
        self.assertFalse(self.module.is_auto_sync_time(datetime(2026, 8, 24, 15, 1)))
        self.assertFalse(self.module.is_auto_sync_time(datetime(2026, 8, 23, 10, 0)))

    def test_auto_sync_deduplicates_all_users_stocks(self):
        now = "2026-08-24T09:00:00"
        with self.module.db_connect() as db:
            first_user = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            second_user = db.execute("INSERT INTO users (name, created_at) VALUES ('auto-second', ?)", (now,)).lastrowid
            for user_id, stock_code in ((first_user, "002156"), (second_user, "002156"), (second_user, "300308")):
                db.execute(
                    """INSERT INTO watchlist_stocks
                    (user_id, stock_code, stock_name, priority, note, created_at, updated_at)
                    VALUES (?, ?, ?, 2, '', ?, ?)""",
                    (user_id, stock_code, stock_code, now, now),
                )

        with patch.object(self.module, "sync_realtime_codes", return_value=(2, [])) as sync:
            result = self.module.run_auto_realtime_sync(datetime(2026, 8, 24, 10, 0))

        self.assertTrue(result)
        sync.assert_called_once_with(["002156", "300308"])
        status = self.module.auto_sync_status()
        self.assertEqual(status["last_success_at"], "2026-08-24T10:00:00")
        self.assertEqual(status["last_synced"], 2)
        self.assertIsNone(status["last_error"])

    def test_auto_sync_skips_outside_window_and_empty_watchlist(self):
        with patch.object(self.module, "sync_realtime_codes") as sync:
            self.assertFalse(self.module.run_auto_realtime_sync(datetime(2026, 8, 24, 12, 0)))
            self.assertFalse(self.module.run_auto_realtime_sync(datetime(2026, 8, 23, 10, 0)))
            self.assertFalse(self.module.run_auto_realtime_sync(datetime(2026, 8, 24, 10, 0)))
        sync.assert_not_called()

    def test_watchlist_kline_only_reads_saved_history(self):
        with patch.object(self.module, "sync_history_codes", return_value=(1, [])):
            self.client.post(
                "/analysis/watchlist",
                data={"stock_code": "300308", "stock_name": "中际旭创", "priority": "2"},
            )
        with self.module.db_connect() as db:
            db.execute(
                """INSERT INTO daily_prices
                (stock_code, trade_date, open, high, low, close, volume, source, fetched_at)
                VALUES ('300308', '2026-08-20', 10, 12, 9, 11, 1000, 'test', '2026-08-21T15:00:00')"""
            )
        with patch.object(self.module, "fetch_daily_prices") as fetch:
            response = self.client.get("/analysis/watchlist/kline/300308")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["data"][0]["date"], "2026-08-20")
        fetch.assert_not_called()

    def test_strategy_filters_and_backtest(self):
        self.client.post("/admin/preview", data={"statement": (self.make_statement(), "statement.xlsx")}, content_type="multipart/form-data")
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW' ORDER BY created_at DESC, rowid DESC LIMIT 1").fetchone()
        self.client.post(f"/admin/import/{job['id']}")

        strategy = self.client.get("/analysis/strategy?min_days=2&max_days=2")
        self.assertEqual(strategy.status_code, 200)
        text = strategy.get_data(as_text=True)
        self.assertIn("持仓 T+2 至 T+2", text)
        self.assertIn("命中样本", text)
        self.assertIn("1", text)
        self.assertIn("002156", text)

        invalid = self.client.get("/analysis/strategy?buy_session=invalid&min_days=9&max_days=2&period=17")
        self.assertEqual(invalid.status_code, 200)
        invalid_text = invalid.get_data(as_text=True)
        self.assertIn("持仓 T+2 至 T+9", invalid_text)
        self.assertIn('<option value="0" selected>全部历史</option>', invalid_text)
        self.assertIn('<option value="all">不限</option>', invalid_text)

    def test_local_users_isolate_imports_and_fifo(self):
        first_user = self.client.get("/admin/statements")
        self.assertIn("yutaoGS", first_user.get_data(as_text=True))
        self.client.post("/users/create", data={"name": "second", "next": "/admin/statements"})
        with self.module.db_connect() as db:
            first_id = db.execute("SELECT id FROM users WHERE name = 'yutaoGS'").fetchone()["id"]
            second_id = db.execute("SELECT id FROM users WHERE name = 'second'").fetchone()["id"]

        self.client.post("/users/switch", data={"user_id": first_id, "next": "/admin/statements"})
        self.client.post("/admin/preview", data={"statement": (self.make_statement(), "first.xlsx")}, content_type="multipart/form-data")
        with self.module.db_connect() as db:
            first_job = db.execute("SELECT id FROM import_jobs WHERE user_id = ?", (first_id,)).fetchone()["id"]
        self.client.post(f"/admin/import/{first_job}")

        self.client.post("/users/switch", data={"user_id": second_id, "next": "/admin/statements"})
        second_page = self.client.get("/admin/statements").get_data(as_text=True)
        self.assertIn("共 0 条成交记录", second_page)
        self.assertEqual(self.client.post(f"/admin/import/{first_job}").status_code, 404)
        preview = self.client.post("/admin/preview", data={"statement": (self.make_statement(), "second.xlsx")}, content_type="multipart/form-data")
        preview_text = preview.get_data(as_text=True)
        self.assertIn("数据库重复", preview_text)
        self.assertIn("预计新增", preview_text)
        with self.module.db_connect() as db:
            second_job = db.execute("SELECT id FROM import_jobs WHERE user_id = ?", (second_id,)).fetchone()["id"]
        self.client.post(f"/admin/import/{second_job}")

        with self.module.db_connect() as db:
            self.assertEqual(db.execute("SELECT COUNT(*) FROM executions WHERE user_id = ?", (first_id,)).fetchone()[0], 3)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM executions WHERE user_id = ?", (second_id,)).fetchone()[0], 3)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM fifo_matches WHERE user_id = ?", (first_id,)).fetchone()[0], 2)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM fifo_matches WHERE user_id = ?", (second_id,)).fetchone()[0], 2)
            self.assertEqual(db.execute("PRAGMA foreign_key_check").fetchall(), [])

    def test_user_management_validation(self):
        response = self.client.post("/users/create", data={"name": "yutaogs", "next": "https://example.com"})
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.headers["Location"], "/analysis/stocks")
        self.assertIn("用户名称已存在", self.client.get("/analysis/stocks").get_data(as_text=True))

        response = self.client.post("/users/create", data={"name": "second", "next": "/analysis/summary"})
        self.assertEqual(response.headers["Location"], "/analysis/summary")
        response = self.client.post("/users/999999/rename", data={"name": "missing", "next": "/analysis/summary"}, follow_redirects=True)
        self.assertIn("用户不存在", response.get_data(as_text=True))

    def test_import_errors_are_visible(self):
        response = self.client.post("/admin/preview", data={}, follow_redirects=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("请选择交割单文件", response.get_data(as_text=True))

        response = self.client.post(
            "/admin/preview",
            data={"statement": (io.BytesIO(b"not a statement"), "statement.pdf")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn("仅支持 .xlsx、.xls、.csv、.txt 文件", response.get_data(as_text=True))

    def test_invalid_trade_error_shows_values_read_from_file(self):
        rows = [
            ["成交日期", "成交时间", "证券代码", "证券名称", "操作", "成交数量", "成交价格"],
            [20260820, "13:19:01", 300209, "行云科技", "证券买入", 300, ""],
        ]
        parsed, errors, ignored, total_rows = self.module.parse_rows(rows)
        self.assertEqual(total_rows, 1)
        self.assertEqual(parsed, [])
        self.assertEqual(ignored, [])
        self.assertIn("数量=300", errors[0])
        self.assertIn("价格=", errors[0])
        self.assertIn("操作=证券买入", errors[0])

    def test_non_trade_rows_are_ignored_instead_of_failed(self):
        rows = [
            ["成交日期", "成交时间", "证券代码", "证券名称", "操作", "成交数量", "成交价格"],
            [20260819, "19:42:00", 301697, "贝特利", "申购配号", 2, 0],
            [20260818, "09:00:00", 600584, "长电科技", "股息入账", 0, 42.82],
        ]
        parsed, errors, ignored, total_rows = self.module.parse_rows(rows)
        self.assertEqual(total_rows, 2)
        self.assertEqual(parsed, [])
        self.assertEqual(errors, [])
        self.assertEqual(len(ignored), 2)
        self.assertIn("申购配号", ignored[0])

    def test_stock_detail(self):
        self.client.post("/admin/preview", data={"statement": (self.make_statement(), "statement.xlsx")}, content_type="multipart/form-data")
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW' ORDER BY created_at DESC, rowid DESC LIMIT 1").fetchone()
        self.client.post(f"/admin/import/{job['id']}")
        self.client.post(
            "/admin/manual",
            data={
                "trade_date": "2026-08-04", "trade_time": "10:00:00",
                "stock_code": "002156", "stock_name": "通富微电", "action": "SELL",
                "quantity": "50", "deal_price": "14", "commission": "0",
                "stamp_tax": "0", "transfer_fee": "0",
            },
        )

        stocks = self.client.get("/analysis/stocks")
        self.assertEqual(stocks.status_code, 200)
        text = stocks.get_data(as_text=True)
        self.assertIn("002156", text)
        self.assertIn("通富微电", text)
        self.assertIn("2026-08-04", text)
        self.assertIn("595.00", text)

        detail = self.client.get("/analysis/stocks?code=002156")
        self.assertEqual(detail.status_code, 200)
        detail_text = detail.get_data(as_text=True)
        self.assertIn("08-01 09:31", detail_text)
        self.assertIn("08-04 10:00", detail_text)
        self.assertIn("总交易笔数", detail_text)
        self.assertIn("交易明细", detail_text)
        self.assertNotIn("<svg", detail_text)
        self.assertNotIn("查看复盘", detail_text)
        self.assertNotIn("FIFO 成本匹配明细", detail_text)

    def test_stock_list_detailed_cycle_metrics(self):
        self._import_workbook_rows([
            [20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
            [20260803, "10:00:00", 2156, "通富微电", "证券卖出", -100, "S1", 13],
        ])
        self._import_workbook_rows([
            [20260805, "09:30:00", 2156, "通富微电", "证券买入", 100, "B2", 10],
            [20260806, "10:00:00", 2156, "通富微电", "证券卖出", -100, "S2", 8],
        ])

        with self.module.db_connect() as db:
            stock = self.module.analysis_data(db)["by_stock"][0]
        self.assertEqual(stock["trades"], 2)
        self.assertAlmostEqual(stock["profit"], 100.0)
        self.assertAlmostEqual(stock["profit_rate"], 0.05)
        self.assertAlmostEqual(stock["win_rate"], 0.5)
        self.assertAlmostEqual(stock["max_profit"], 300.0)
        self.assertAlmostEqual(stock["max_loss"], -200.0)
        self.assertAlmostEqual(stock["avg_holding_days"], 1.5)

        page = self.client.get("/analysis/stocks").get_data(as_text=True)
        for label in ("累计盈亏", "累计收益率", "交易总次数", "胜率", "最大盈利", "最大亏损", "平均持仓天数"):
            self.assertIn(label, page)
        self.assertIn("+100.00 元", page)
        self.assertIn("+5.0%", page)
        self.assertIn("2 笔", page)
        self.assertIn("50%", page)
        self.assertIn("+300.00", page)
        self.assertIn("-200.00", page)
        self.assertIn("1.5 天", page)

        self.assertIn("总交易笔数", page)
        self.assertIn("盈亏比", page)
        self.assertIn("1.50", page)
        self.assertIn("+100.00", page)
        self.assertIn("+30.00%", page)
        self.assertIn("-20.00%", page)
        self.assertLess(page.index("08-06 10:00"), page.index("08-03 10:00"))

    def test_stock_detail_same_day_and_no_loss_profit_factor(self):
        self._import_workbook_rows([
            [20260801, "09:30:00", 2156, "通富微电", "证券买入", 100, "B1", 10],
            [20260801, "14:30:00", 2156, "通富微电", "证券卖出", -100, "S1", 11],
        ])
        page = self.client.get("/analysis/stocks?code=002156").get_data(as_text=True)
        self.assertIn("08-01 09:30", page)
        self.assertIn("08-01 14:30", page)
        self.assertIn("&lt;1天", page)
        self.assertIn("--", page)

    def test_stock_list_sorting(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["成交日期", "成交时间", "证券代码", "证券名称", "操作", "成交数量", "成交编号", "成交价格", "成交金额"])
        sheet.append([20260801, "09:30:00", 100001, "较早盈利", "证券买入", 100, "A1", 10, 1000])
        sheet.append([20260802, "10:00:00", 100001, "较早盈利", "证券卖出", -100, "A2", 20, 2000])
        sheet.append([20260810, "09:30:00", 100002, "最近亏损", "证券买入", 100, "B1", 20, 2000])
        sheet.append([20260811, "10:00:00", 100002, "最近亏损", "证券卖出", -100, "B2", 15, 1500])
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        self.client.post("/admin/preview", data={"statement": (output, "sorting.xlsx")}, content_type="multipart/form-data")
        with self.module.db_connect() as db:
            job = db.execute("SELECT id FROM import_jobs WHERE status = 'PREVIEW'").fetchone()
        self.client.post(f"/admin/import/{job['id']}")

        recent = self.client.get("/analysis/stocks").get_data(as_text=True)
        self.assertLess(recent.index("最近亏损"), recent.index("较早盈利"))
        self.assertIn('<option value="recent" selected>最近交易</option>', recent)
        self.assertIn("2026-08-11", recent)

        profit = self.client.get("/analysis/stocks?sort=profit").get_data(as_text=True)
        self.assertLess(profit.index("较早盈利"), profit.index("最近亏损"))
        self.assertIn('<option value="profit" selected>盈利最高</option>', profit)
        self.assertIn("sort=profit", profit)

        loss = self.client.get("/analysis/stocks?sort=loss").get_data(as_text=True)
        self.assertLess(loss.index("最近亏损"), loss.index("较早盈利"))

        invalid = self.client.get("/analysis/stocks?sort=invalid").get_data(as_text=True)
        self.assertLess(invalid.index("最近亏损"), invalid.index("较早盈利"))
        self.assertIn('<option value="recent" selected>最近交易</option>', invalid)

        by_name = self.client.get("/analysis/stocks?q=较早&sort=profit").get_data(as_text=True)
        self.assertIn("较早盈利", by_name)
        self.assertNotIn("最近亏损", by_name)
        self.assertIn("找到 1 只股票", by_name)
        self.assertIn("q=%E8%BE%83%E6%97%A9", by_name)
        self.assertIn("sort=profit", by_name)

        by_code = self.client.get("/analysis/stocks?q=100002").get_data(as_text=True)
        self.assertIn("最近亏损", by_code)
        self.assertNotIn("较早盈利", by_code)

        partial_code = self.client.get("/analysis/stocks?q=0002").get_data(as_text=True)
        self.assertIn("最近亏损", partial_code)

        empty = self.client.get("/analysis/stocks?q=不存在").get_data(as_text=True)
        self.assertIn("没有匹配名称或代码的股票", empty)


    def test_hot_sectors_page_renders(self):
        response = self.client.get("/analysis/hot-sectors")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("热门板块", html)
        self.assertIn("hot-sectors-grid", html)
        self.assertIn("hot-sectors-refresh", html)
        self.assertIn("/api/hot-sectors", html)


    def test_hot_sectors_df_records_conversion(self):
        frame = pd.DataFrame([{
            "排名": 1, "板块名称": "测试板块", "最新价": 9.9,
            "涨跌幅": float("nan"), "上涨家数": 3, "领涨股票": None,
        }])
        records = self.module.df_records(frame, limit=1)
        self.assertEqual(records[0]["排名"], 1)
        self.assertEqual(records[0]["板块名称"], "测试板块")
        self.assertEqual(records[0]["涨跌幅"], None)
        self.assertEqual(records[0]["上涨家数"], 3)
        self.assertIsNone(records[0]["领涨股票"])
        self.assertEqual(self.module.df_records(None), [])
        self.assertEqual(self.module.df_records(pd.DataFrame()), [])


    def test_hot_sectors_normalize_board_frame(self):
        frame = pd.DataFrame([
            {"板块": "甲板块", "label": "new_abc", "平均价格": 10.0, "涨跌额": 0.5, "涨跌幅": 5.0,
             "公司家数": 20, "总成交额": 3.2e9, "股票名称": "A股票", "个股-涨跌幅": 10.0},
            {"板块": "乙板块", "label": "new_def", "平均价格": 5.0, "涨跌额": -0.2, "涨跌幅": -2.0,
             "公司家数": 8, "总成交额": 1.1e8, "股票名称": "B股票", "个股-涨跌幅": 3.0},
        ])
        normalized = self.module.normalize_board_frame(frame)
        self.assertEqual(list(normalized["板块名称"]), ["甲板块", "乙板块"])
        self.assertEqual(normalized.loc[0, "板块代码"], "new_abc")
        self.assertEqual(normalized.loc[0, "领涨股票"], "A股票")
        self.assertEqual(normalized.loc[0, "成交额"], 3.2e9)
        self.assertEqual(normalized.loc[0, "领涨股票-涨跌幅"], 10.0)
        self.assertTrue(self.module.normalize_board_frame(pd.DataFrame()).empty)


    def test_hot_sectors_build_hot_rank_records(self):
        rank_rows = [
            {"rank": 1, "code": "688836", "name": "sh688836"},
            {"rank": 2, "code": "002716", "name": "sz002716"},
            {"rank": 3, "code": "999999", "name": "sz999999"},
        ]
        quotes = {
            "688836": {"close": 603.08, "previous_close": 672.41},
            "002716": {"close": 11.46, "previous_close": 10.42},
        }
        records = self.module.build_hot_rank_records(rank_rows, quotes)
        self.assertEqual(len(records), 3)
        self.assertEqual(records[0]["排名"], 1)
        self.assertEqual(records[0]["代码"], "688836")
        self.assertAlmostEqual(records[0]["涨跌幅"], -10.3107, places=3)
        self.assertAlmostEqual(records[1]["涨跌额"], 1.04, places=2)
        self.assertIsNone(records[2]["最新价"])
        self.assertIsNone(records[2]["涨跌幅"])


    def test_hot_sectors_api_returns_three_modules(self):
        industry = pd.DataFrame([{
            "排名": 1, "板块名称": "小金属", "板块代码": 881101, "最新价": 1200.5,
            "涨跌额": 45.0, "涨跌幅": 3.9, "总市值": 5.2e11, "换手率": 2.1,
            "上涨家数": 20, "下跌家数": 3, "领涨股票": "X科技", "领涨股票-涨跌幅": 10.0,
        }])
        concept = pd.DataFrame([{
            "排名": 1, "板块名称": "可控核聚变", "板块代码": 881202, "最新价": 888.8,
            "涨跌额": -5.5, "涨跌幅": -0.6, "总市值": float("nan"), "换手率": 1.2,
            "上涨家数": 5, "下跌家数": 9, "领涨股票": None, "领涨股票-涨跌幅": None,
        }])
        rank = pd.DataFrame([{
            "当前排名": 1, "代码": "SZ000665", "股票名称": "湖北广电",
            "最新价": 5.6, "涨跌额": 0.31, "涨跌幅": 5.9,
        }])
        with patch.object(self.module, "fetch_hot_industry", return_value=industry), \
             patch.object(self.module, "fetch_hot_concept", return_value=concept), \
             patch.object(self.module, "fetch_hot_rank", return_value=rank):
            response = self.client.get("/api/hot-sectors?force=1")
        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        modules = payload["modules"]
        self.assertEqual(set(modules), {"industry", "concept", "hot_rank"})
        self.assertEqual(modules["industry"]["records"][0]["板块代码"], 881101)
        self.assertIsNone(modules["concept"]["records"][0]["总市值"])
        self.assertEqual(modules["hot_rank"]["records"][0]["代码"], "SZ000665")
        self.assertIsNone(modules["industry"]["error"])
        self.assertIsNone(modules["concept"]["error"])
        self.assertIsNone(modules["hot_rank"]["error"])
        self.assertIsNotNone(payload["updated_at"])


    def test_hot_sectors_partial_failure_keeps_last_good(self):
        industry = pd.DataFrame([{"板块名称": "小金属", "涨跌幅": 3.9}])
        rank = pd.DataFrame([{"代码": "SZ000665", "股票名称": "湖北广电", "最新价": 5.6, "涨跌幅": 5.9}])
        with patch.object(self.module, "fetch_hot_industry", return_value=industry), \
             patch.object(self.module, "fetch_hot_concept", return_value=industry), \
             patch.object(self.module, "fetch_hot_rank", return_value=rank):
            self.client.get("/api/hot-sectors?force=1")
        with patch.object(self.module, "fetch_hot_industry", return_value=industry), \
             patch.object(self.module, "fetch_hot_concept", return_value=industry), \
             patch.object(self.module, "fetch_hot_rank", side_effect=RuntimeError("接口异常")):
            payload = self.module.hot_sectors_data(force=True)
        self.assertEqual(len(payload["modules"]["hot_rank"]["records"]), 1)
        self.assertIn("人气榜：", payload["modules"]["hot_rank"]["error"])
        self.assertEqual(len(payload["modules"]["industry"]["records"]), 1)
        self.assertIsNone(payload["modules"]["industry"]["error"])


    def test_hot_sectors_first_failure_surfaces_error(self):
        self.module.HOT_SECTORS_CACHE.update({"modules": {}, "updated_at": None})
        with patch.object(self.module, "fetch_hot_industry", side_effect=RuntimeError("网络错误")), \
             patch.object(self.module, "fetch_hot_concept", return_value=pd.DataFrame()), \
             patch.object(self.module, "fetch_hot_rank", side_effect=RuntimeError("接口错误")):
            payload = self.module.hot_sectors_data(force=True)
        self.assertEqual(payload["modules"]["industry"]["records"], [])
        self.assertIn("行业板块：", payload["modules"]["industry"]["error"])
        self.assertEqual(payload["modules"]["concept"]["records"], [])
        self.assertIsNone(payload["modules"]["concept"]["error"])
        self.assertEqual(payload["modules"]["hot_rank"]["records"], [])
        self.assertIn("人气榜：", payload["modules"]["hot_rank"]["error"])


if __name__ == "__main__":
    unittest.main()
