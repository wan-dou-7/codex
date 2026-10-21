from decimal import Decimal
import os
import unittest

from store_recon import RosterEntry, StoreSpend, reconcile


class ReconcileTests(unittest.TestCase):
    def test_budget_status_edges_and_unconfirmed_name_is_not_auto_matched(self):
        stores = {
            "10001": StoreSpend("10001", "示例瑜伽馆甲", "杭州市", Decimal("240.00")),
            "10002": StoreSpend("10002", "示例普拉提馆乙", "宁波市", Decimal("300.01")),
            "10003": StoreSpend("10003", "示例瑜伽馆丙", "苏州市", Decimal("239.99")),
        }
        entries = [
            RosterEntry("示例瑜伽馆甲", Decimal("300"), "杭州市"),
            RosterEntry("示例普拉提馆乙", Decimal("300"), "宁波市"),
            RosterEntry("示例瑜伽馆丙", Decimal("300"), "苏州市"),
            RosterEntry("示例瑜伽馆丁", Decimal("900"), "无锡市"),
        ]
        report = reconcile(entries, stores)
        self.assertEqual([r["状态"] for r in report["results"]], ["临近预算", "已超预算", "正常"])
        self.assertEqual(len(report["pending"]), 1)


    def test_city_disambiguates_same_name_and_manual_mapping_uses_store_id(self):
        stores = {
            "20001": StoreSpend("20001", "伽年华示例馆", "杭州市", Decimal("80")),
            "20002": StoreSpend("20002", "伽年华示例馆", "宁波市", Decimal("90")),
        }
        report = reconcile([RosterEntry("伽年华示例馆", Decimal("300"), "宁波市")], stores)
        self.assertEqual(report["results"][0]["门店ID"], "20002")
        report = reconcile([RosterEntry("示例馆新别名", Decimal("300"), "杭州市")], stores,
                           {"示例馆新别名": "20001"})
        self.assertEqual(report["results"][0]["门店ID"], "20001")
        self.assertEqual(report["results"][0]["匹配方式"], "人工映射")


    def test_ambiguous_same_name_without_city_stays_pending(self):
        stores = {
            "30001": StoreSpend("30001", "同名示例店", "杭州市", Decimal("10")),
            "30002": StoreSpend("30002", "同名示例店", "宁波市", Decimal("20")),
        }
        report = reconcile([RosterEntry("同名示例店", Decimal("600"))], stores)
        self.assertFalse(report["results"])
        self.assertEqual(len(report["pending"]), 1)
        self.assertIn("不唯一", report["pending"][0]["原因"])


    @unittest.skipUnless(os.getenv("HISTORICAL_REFERENCE_XLSX"), "本地未提供历史核验文件")
    def test_historical_october_15_reference_if_supplied(self):
        """The source workbook stays private; set this env var locally to verify it."""
        from pathlib import Path
        reference = os.getenv("HISTORICAL_REFERENCE_XLSX")
        from store_recon import load_roster, load_spend

        path = Path(reference)
        entries, _ = load_roster([path], ["激励名单"])
        stores = load_spend([path], ["投流明细"])
        report = reconcile(entries, stores)
        self.assertEqual(len(report["results"]), 177)
        self.assertEqual(len(report["pending"]), 11)
        self.assertEqual(len(report["overspend"]), 43)
        self.assertEqual(round(sum(r["实际消耗（元）"] for r in report["results"]), 2), 75803.40)
        self.assertEqual(round(sum(-r["剩余/超支（元）"] for r in report["overspend"]), 2), 16711.28)


if __name__ == "__main__":
    unittest.main()
