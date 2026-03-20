import json
import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from aliexpress_export import (
    AliExpressMtopClient,
    BrowserProfile,
    OrderBundle,
    OrderStore,
    coerce_order_date,
    detect_installed_browsers,
    find_browser_app_path,
    firefox_profile_fallback_label,
    load_chromium_profile_labels,
    load_firefox_profile_labels,
    parse_optional_int,
    repeated_pagination_reason,
    safe_filename,
    try_parse_date_from_text,
    uniquify_profile_labels,
)


class HelpersTest(unittest.TestCase):
    def test_try_parse_date_from_text_supports_multiple_formats(self) -> None:
        self.assertEqual(try_parse_date_from_text("Created on 2025-01-15"), date(2025, 1, 15))
        self.assertEqual(try_parse_date_from_text("Jan 5, 2024"), date(2024, 1, 5))
        self.assertEqual(try_parse_date_from_text("05 Feb 2023"), date(2023, 2, 5))

    def test_coerce_order_date_supports_epoch_millis(self) -> None:
        self.assertEqual(coerce_order_date(1736035200000), date(2025, 1, 5))

    def test_safe_filename_strips_problematic_characters(self) -> None:
        self.assertEqual(safe_filename("invoice: 123/456"), "invoice_123_456")

    def test_order_store_filters_by_date_range(self) -> None:
        store = OrderStore()
        store.orders["1"] = OrderBundle(order_id="1", list_fields={"orderDateText": "2025-01-10"})
        store.orders["2"] = OrderBundle(order_id="2", list_fields={"orderDateText": "2025-02-10"})
        results = store.filtered(date(2025, 1, 1), date(2025, 1, 31))
        self.assertEqual([item.order_id for item in results], ["1"])

    def test_load_chromium_profile_labels_reads_local_state(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            support_dir = Path(tmp_dir)
            local_state = (
                '{"profile":{"info_cache":{"Default":{"name":"Work"},'
                '"Profile 1":{"name":"Personal"}}}}'
            )
            (support_dir / "Local State").write_text(
                local_state,
                encoding="utf-8",
            )
            self.assertEqual(
                load_chromium_profile_labels(support_dir),
                {"Default": "Work", "Profile 1": "Personal"},
            )

    def test_load_firefox_profile_labels_reads_profiles_ini(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            profiles_root = Path(tmp_dir) / "Firefox" / "Profiles"
            profiles_root.mkdir(parents=True)
            profiles_ini = profiles_root.parent / "profiles.ini"
            profiles_ini.write_text(
                "[Profile0]\nName=Default Release\nPath=Profiles/abcd.default-release\n",
                encoding="utf-8",
            )
            self.assertEqual(
                load_firefox_profile_labels(profiles_root),
                {"abcd.default-release": "Default Release"},
            )

    def test_parse_optional_int_supports_strings(self) -> None:
        self.assertEqual(parse_optional_int("12"), 12)
        self.assertIsNone(parse_optional_int("12a"))

    def test_firefox_profile_fallback_label_strips_random_prefix(self) -> None:
        self.assertEqual(
            firefox_profile_fallback_label("OyLS3fOI.Profile 1"),
            "Profile 1",
        )
        self.assertEqual(
            firefox_profile_fallback_label("jfv5yhgw.default-release"),
            "default-release",
        )

    def test_uniquify_profile_labels_only_adds_suffix_on_duplicates(self) -> None:
        profiles = [
            BrowserProfile(
                browser_id="firefox",
                browser_label="Firefox",
                profile_label="default-release",
                profile_path=Path("/tmp/a.default-release"),
                cookie_file=Path("/tmp/a.default-release/cookies.sqlite"),
            ),
            BrowserProfile(
                browser_id="firefox",
                browser_label="Firefox",
                profile_label="default-release",
                profile_path=Path("/tmp/b.Profile 2"),
                cookie_file=Path("/tmp/b.Profile 2/cookies.sqlite"),
            ),
            BrowserProfile(
                browser_id="firefox",
                browser_label="Firefox",
                profile_label="Profile 1",
                profile_path=Path("/tmp/c.Profile 1"),
                cookie_file=Path("/tmp/c.Profile 1/cookies.sqlite"),
            ),
        ]
        result = uniquify_profile_labels(profiles)
        self.assertEqual(
            [profile.profile_label for profile in result],
            [
                "default-release (a.default-release)",
                "default-release (b.Profile 2)",
                "Profile 1",
            ],
        )

    def test_find_browser_app_path_prefers_known_app_paths(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            app_path = Path(tmp_dir) / "Firefox.app"
            app_path.mkdir()
            meta = {"app_paths": [str(app_path)], "bundle_id": "org.mozilla.firefox"}
            self.assertEqual(find_browser_app_path(meta), app_path)

    def test_find_browser_app_path_falls_back_to_spotlight(self) -> None:
        app_path = Path("/Applications/Fake Browser.app")
        meta = {"app_paths": [], "bundle_id": "com.example.fake"}
        with patch(
            "aliexpress_export.spotlight_app_paths",
            return_value=[app_path],
        ):
            self.assertEqual(find_browser_app_path(meta), app_path)

    def test_detect_installed_browsers_ignores_support_dirs_without_app(self) -> None:
        fake_browsers = {
            "installed": {
                "label": "Installed",
                "app_paths": [],
                "bundle_id": "com.example.installed",
                "support_dir": "/tmp/installed",
                "family": "chromium",
            },
            "stale": {
                "label": "Stale",
                "app_paths": [],
                "bundle_id": "com.example.stale",
                "support_dir": "/tmp/stale",
                "family": "chromium",
            },
        }
        with patch("aliexpress_export.SUPPORTED_BROWSERS", fake_browsers):
            with patch(
                "aliexpress_export.find_browser_app_path",
                side_effect=[Path("/Applications/Installed.app"), None],
            ):
                self.assertEqual(detect_installed_browsers(), ["installed"])

    def test_fetch_order_list_page_more_uses_ultron_post_shape(self) -> None:
        previous_payload = {
            "data": {
                "data": {
                    "pc_om_list_body_1": {
                        "fields": {"hasMore": True, "pageIndex": 1, "pageSize": 10},
                        "tag": "pc_om_list_body",
                    },
                    "pc_om_list_header_action_2": {
                        "fields": {
                            "searchInput": "",
                            "searchOption": "order",
                            "statusTab": "all",
                            "timeOption": "all",
                        },
                        "tag": "pc_om_list_header_action",
                    },
                },
                "linkage": {"signature": "abc"},
                "hierarchy": {"structure": {"pc_om_list_body_1": ["pc_om_list_order_1"]}},
                "endpoint": {"droplet": True},
            }
        }
        client = AliExpressMtopClient(
            [
                {
                    "name": "_m_h5_tk",
                    "value": "token_123456789",
                    "domain": ".aliexpress.com",
                    "path": "/",
                }
            ],
            ship_to_country="UK",
            lang="en_US",
        )
        try:
            with patch.object(client, "_request", return_value={}) as request_mock:
                client.fetch_order_list_page_more(previous_payload, 2, 10)
            request_mock.assert_called_once()
            call_args = request_mock.call_args
            self.assertEqual(call_args.args[0], "mtop.aliexpress.trade.buyer.order.list")
            payload = call_args.args[1]
            params = json.loads(payload["params"])
            request_blocks = json.loads(params["data"])
            self.assertEqual(request_blocks["pc_om_list_body_1"]["fields"]["pageIndex"], 2)
            self.assertEqual(request_blocks["pc_om_list_body_1"]["fields"]["pageSize"], 10)
            self.assertEqual(params["operator"], "pc_om_list_body_1")
            self.assertEqual(payload["shipToCountry"], "UK")
            self.assertEqual(payload["_lang"], "en_US")
            self.assertEqual(call_args.kwargs["method"], "POST")
            self.assertEqual(call_args.kwargs["request_type"], "originaljson")
            self.assertEqual(call_args.kwargs["extra_query_params"], {"post": "1", "isSec": "1"})
        finally:
            client.close()

    def test_repeated_pagination_reason_stops_on_duplicate_page(self) -> None:
        self.assertEqual(
            repeated_pagination_reason(
                requested_page_index=2,
                new_orders=0,
                reported_page_index=2,
                last_reported_page_index=1,
            ),
            "AliExpress returned no new orders.",
        )

    def test_repeated_pagination_reason_stops_on_non_advancing_page_index(self) -> None:
        self.assertEqual(
            repeated_pagination_reason(
                requested_page_index=3,
                new_orders=2,
                reported_page_index=2,
                last_reported_page_index=2,
            ),
            "AliExpress reported pageIndex=2 again.",
        )


if __name__ == "__main__":
    unittest.main()
