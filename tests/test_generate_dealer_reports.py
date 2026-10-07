from __future__ import annotations

import unittest

from scripts.generate_dealer_reports import (
    CrawlSummary,
    EndpointRecord,
    FormRecord,
    ObjectRecord,
    ReportBundle,
    ScrapeRecommendation,
    classify_url_family,
    detect_tech_flags,
    normalize_url,
    parse_robots_txt,
    parse_sitemap_xml,
    render_report,
    slugify,
)


class GenerateDealerReportsTests(unittest.TestCase):
    def test_slugify(self) -> None:
        self.assertEqual(slugify("KServico Trade, Inc."), "kservico-trade-inc")
        self.assertEqual(slugify("  Premio Corporation  "), "premio-corporation")

    def test_normalize_url(self) -> None:
        self.assertEqual(
            normalize_url("http://Example.com/path//to/page/#section"),
            "https://example.com/path/to/page",
        )
        self.assertEqual(normalize_url("https://example.com/foo/"), "https://example.com/foo")

    def test_parse_robots_txt(self) -> None:
        text = """
        User-agent: *
        Disallow: /private/
        Sitemap: https://example.com/sitemap.xml
        Sitemap: https://example.com/wp-sitemap.xml
        """
        disallows, sitemaps = parse_robots_txt(text)
        self.assertEqual(disallows, ["/private/"])
        self.assertEqual(
            sitemaps,
            ["https://example.com/sitemap.xml", "https://example.com/wp-sitemap.xml"],
        )

    def test_parse_sitemap_urlset(self) -> None:
        xml = b"""<?xml version="1.0" encoding="UTF-8"?>
        <urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
          <url><loc>https://example.com/</loc></url>
          <url><loc>https://example.com/store</loc></url>
        </urlset>"""
        kind, urls = parse_sitemap_xml(xml)
        self.assertEqual(kind, "urlset")
        self.assertEqual(urls, ["https://example.com/", "https://example.com/store"])

    def test_parse_sitemap_index(self) -> None:
        xml = b"""<?xml version="1.0" encoding="UTF-8"?>
        <sitemapindex xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
          <sitemap><loc>https://example.com/sitemap-posts.xml</loc></sitemap>
          <sitemap><loc>https://example.com/sitemap-pages.xml</loc></sitemap>
        </sitemapindex>"""
        kind, urls = parse_sitemap_xml(xml)
        self.assertEqual(kind, "sitemapindex")
        self.assertEqual(
            urls,
            ["https://example.com/sitemap-posts.xml", "https://example.com/sitemap-pages.xml"],
        )

    def test_classify_url_family(self) -> None:
        self.assertEqual(classify_url_family("https://example.com/"), "homepage")
        self.assertEqual(
            classify_url_family("https://example.com/branch-locator/"),
            "branches / store locator",
        )
        self.assertEqual(
            classify_url_family("https://example.com/motorcycles/honda-click"),
            "products / vehicles",
        )
        self.assertEqual(
            classify_url_family("https://example.com/lending-application/"),
            "forms / applications / calculators",
        )

    def test_detect_tech_flags(self) -> None:
        html = """
        <html>
          <script src="/_next/static/chunk.js"></script>
          <script>var ajaxUrl = '/wp-admin/admin-ajax.php';</script>
          <div class="agile-store-locator"></div>
        </html>
        """
        headers = {"server": "cloudflare", "x-powered-by": "Next.js"}
        flags = detect_tech_flags(html, "https://example.com", headers, ["/_next/static/chunk.js"])
        self.assertIn("nextjs", flags)
        self.assertIn("cloudflare", flags)
        self.assertIn("admin-ajax", flags)
        self.assertIn("agile-store-locator", flags)

    def test_render_report_contains_required_sections(self) -> None:
        bundle = ReportBundle(
            dealer="Example Dealer",
            website="https://example.com",
            crawl_summary=CrawlSummary(
                homepage_status=200,
                final_homepage_url="https://example.com",
                robots_status=200,
                sitemap_urls=["https://example.com/sitemap.xml"],
                discovered_urls=5,
                fetched_pages=5,
                html_pages=5,
                blocked_pages=0,
                errors=0,
            ),
            url_families=[("homepage", 1, ["https://example.com"], "Homepage discovered.")],
            objects=[
                ObjectRecord(
                    kind="branch / store records",
                    source_url="https://example.com/store",
                    fields=["branch_name", "address"],
                    extraction_hint="Rendered store cards expose branch information.",
                    confidence="High",
                )
            ],
            endpoints=[
                EndpointRecord(
                    source_url="https://example.com/store",
                    endpoint="https://example.com/wp-json/demo/v1/stores",
                    kind="WordPress REST",
                    auth_required=False,
                    blocked=False,
                    notes="Detected in markup.",
                    sample_fields=["name", "address"],
                )
            ],
            forms=[
                (
                    "https://example.com/contact",
                    FormRecord(
                        action="https://example.com/contact",
                        method="POST",
                        visible_fields=["name", "email"],
                        hidden_fields=["token"],
                        submit_labels=["Submit"],
                    ),
                )
            ],
            tech_flags=["wordpress", "wp-json"],
            constraints=["No severe blocker."],
            overview_notes=["Uses WordPress."],
            open_questions=["None."],
            scrape_recommendation=ScrapeRecommendation(
                primary_method="Direct public endpoint extraction",
                fallback_method="Rendered DOM parsing",
                tooling="requests",
                blockers="Low",
                confidence="High",
            ),
        )

        report = render_report(bundle)
        for heading in (
            "# Example Dealer",
            "## Overview",
            "## Crawl Coverage",
            "## Site Architecture and Rendering Model",
            "## Discoverable URL Families",
            "## Scrapable Objects and Fields",
            "## APIs, Feeds, and Embedded Data Sources",
            "## Forms and Interactive Flows",
            "## Anti-Bot, Access, and Constraints",
            "## Recommended Scrape Method",
            "## Open Questions and Risks",
        ):
            self.assertIn(heading, report)


if __name__ == "__main__":
    unittest.main()
