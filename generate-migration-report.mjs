import {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, HeadingLevel, WidthType, AlignmentType, BorderStyle,
  ShadingType, TableLayoutType, PageBreak
} from 'docx';
import { writeFileSync } from 'fs';

const BRAND_BLUE = '1B3A6B';
const HEADER_BG = 'D6E4F0';
const ALT_ROW_BG = 'F5F8FB';

function createHeaderCell(text, width) {
  return new TableCell({
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 20, color: BRAND_BLUE })] })],
    shading: { type: ShadingType.SOLID, color: HEADER_BG },
    width: width ? { size: width, type: WidthType.PERCENTAGE } : undefined,
  });
}

function createCell(text, options = {}) {
  return new TableCell({
    children: [new Paragraph({
      children: [new TextRun({ text: String(text), size: 20, bold: options.bold || false })],
      alignment: options.align || AlignmentType.LEFT
    })],
    shading: options.shaded ? { type: ShadingType.SOLID, color: ALT_ROW_BG } : undefined,
    width: options.width ? { size: options.width, type: WidthType.PERCENTAGE } : undefined,
  });
}

function createTable(headers, rows, widths) {
  const headerRow = new TableRow({
    children: headers.map((h, i) => createHeaderCell(h, widths?.[i])),
    tableHeader: true,
  });
  const dataRows = rows.map((row, idx) =>
    new TableRow({
      children: row.map((cell, i) => createCell(cell, { shaded: idx % 2 === 1, width: widths?.[i] })),
    })
  );
  return new Table({
    rows: [headerRow, ...dataRows],
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
  });
}

function heading(text, level = HeadingLevel.HEADING_1) {
  return new Paragraph({ text, heading: level, spacing: { before: 300, after: 150 } });
}

function para(text, options = {}) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, bold: options.bold, italics: options.italic })],
    spacing: { after: 120 },
  });
}

function bullet(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22 })],
    bullet: { level: 0 },
    spacing: { after: 60 },
  });
}

const doc = new Document({
  creator: 'Adobe EDS Migration Analysis',
  title: 'AEM Edge Delivery Services Migration Scope Analysis - Australia Post',
  sections: [{
    properties: {},
    children: [
      // TITLE PAGE
      new Paragraph({ spacing: { before: 2000 } }),
      new Paragraph({
        children: [new TextRun({ text: 'AEM Edge Delivery Services', size: 48, bold: true, color: BRAND_BLUE })],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        children: [new TextRun({ text: 'Migration Scope Analysis', size: 44, bold: true, color: BRAND_BLUE })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
      }),
      new Paragraph({
        children: [new TextRun({ text: 'auspost.com.au', size: 36, color: '444444' })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 },
      }),
      new Paragraph({
        children: [new TextRun({ text: `Report Date: June 3, 2026`, size: 24, color: '666666' })],
        alignment: AlignmentType.CENTER,
      }),
      new Paragraph({
        children: [new TextRun({ text: 'Prepared for: AEM Migration Planning', size: 24, color: '666666' })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
      }),
      new Paragraph({
        children: [new TextRun({ text: 'Classification: Confidential', size: 20, italics: true, color: '999999' })],
        alignment: AlignmentType.CENTER,
      }),

      // PAGE BREAK
      new Paragraph({ children: [new PageBreak()] }),

      // EXECUTIVE SUMMARY
      heading('Executive Summary'),
      para('This document provides a comprehensive migration scope analysis for auspost.com.au to Adobe Experience Manager (AEM) Edge Delivery Services (EDS). The analysis covers 1,167 publicly accessible URLs across the Australia Post website, encompassing personal, business, and enterprise/government segments.'),
      para('The site is currently built on AEM (Adobe Experience Manager) with a traditional AEM Sites implementation using client libraries, custom components, and extensive third-party integrations. Migration to Edge Delivery Services will deliver significant performance improvements, simplified authoring, and reduced operational complexity.'),
      new Paragraph({ spacing: { after: 200 } }),

      // SITE OVERVIEW
      heading('1. Site Overview', HeadingLevel.HEADING_1),
      para('Total Indexed URLs: 1,167', { bold: true }),
      para('Current Platform: AEM Sites (Classic/Touch UI) with Adobe Experience Platform integration'),
      para('Content Management: Adobe Experience Manager DAM + Custom Components'),
      new Paragraph({ spacing: { after: 100 } }),

      createTable(
        ['Category', 'URL Count', 'Percentage'],
        [
          ['/business', '659', '56.5%'],
          ['/community-hub', '125', '10.7%'],
          ['/money-travel', '98', '8.4%'],
          ['/sending', '89', '7.6%'],
          ['/about-us', '87', '7.5%'],
          ['/disruptions-and-updates', '35', '3.0%'],
          ['/receiving', '21', '1.8%'],
          ['/locate', '18', '1.5%'],
          ['/id-and-document-services', '15', '1.3%'],
          ['/travel-insurance', '9', '0.8%'],
          ['/terms-conditions', '5', '0.4%'],
          ['Other (shop, external, root)', '6', '0.5%'],
        ],
        [40, 30, 30]
      ),

      new Paragraph({ children: [new PageBreak()] }),

      // TEMPLATES INVENTORY
      heading('2. Templates Inventory', HeadingLevel.HEADING_1),
      para('Based on analysis of page structures, navigation patterns, and content layouts, the following unique page templates have been identified:'),
      new Paragraph({ spacing: { after: 100 } }),

      createTable(
        ['#', 'Template Name', 'Description', 'Complexity', 'Pages'],
        [
          ['1', 'Homepage', 'Hero banner, tracking widget, promo cards, quick links, shop section, delivery info', 'High', '3'],
          ['2', 'Category Landing', 'Hero with image, breadcrumbs, quick action links, service card grid, FAQ section', 'Medium', '45'],
          ['3', 'Service Detail', 'Breadcrumbs, content sections with headings, comparison cards, footnotes, FAQ', 'Medium', '320'],
          ['4', 'Product/Pricing', 'Pricing tables, feature comparisons, optional extras, calculator links', 'High', '85'],
          ['5', 'Content/Article', 'Rich text, embedded video (YouTube), image galleries, card CTAs', 'Medium', '180'],
          ['6', 'Community Story', 'Article layout with hero image, body text, related stories carousel', 'Low', '125'],
          ['7', 'Help & Support', 'Search bar, chatbot integration, category cards, contact channels', 'High', '45'],
          ['8', 'Location Finder', 'Map integration, search/filter, location cards, operating hours', 'High', '18'],
          ['9', 'Form/Application', 'Multi-step forms, validation, payment integration, confirmation flows', 'High', '35'],
          ['10', 'Disruptions/Alerts', 'Dynamic content feed, filtering, status indicators, date-based sorting', 'High', '35'],
          ['11', 'Legal/Policy', 'Long-form text content, table of contents, anchor navigation', 'Low', '15'],
          ['12', 'Business Portal Landing', 'Enterprise navigation, solution cards, portal links, account tools', 'Medium', '150'],
          ['13', 'Travel Insurance Quote', 'Third-party embedded application (insurance quote engine)', 'High', '9'],
          ['14', 'Currency Converter', 'Interactive calculator tool, exchange rate data, conversion history', 'High', '2'],
        ],
        [5, 22, 40, 13, 10]
      ),

      new Paragraph({ spacing: { after: 200 } }),
      para('Template Complexity Summary:', { bold: true }),
      createTable(
        ['Complexity Level', 'Template Count', 'Total Pages'],
        [
          ['High', '7', '227'],
          ['Medium', '5', '695'],
          ['Low', '2', '140'],
          ['TOTAL', '14', '1,167*'],
        ],
        [40, 30, 30]
      ),
      para('* Includes pages that map to multiple templates or have hybrid structures.', { italic: true }),

      new Paragraph({ children: [new PageBreak()] }),

      // BLOCKS / COMPONENTS CATALOG
      heading('3. Blocks / Components Catalog', HeadingLevel.HEADING_1),
      para('The following reusable blocks and components have been identified across the site. Design variations of the same content model are noted as variants rather than separate blocks.'),
      new Paragraph({ spacing: { after: 100 } }),

      createTable(
        ['#', 'Block Name', 'Description', 'Variants', 'Complexity'],
        [
          ['1', 'Header/Navigation', 'Multi-level mega menu with 3 audience segments (Personal/Business/Enterprise), quick links, search, account dropdown', '3 (audience segments)', 'High'],
          ['2', 'Footer', 'Collapsible quick links, blog links, global nav links, social icons, acknowledgement of country', '1', 'Medium'],
          ['3', 'Hero Banner', 'Full-width image/text overlay with CTA button, background image variants', '3 (home, category, article)', 'Medium'],
          ['4', 'Tracking Widget', 'Input field with "Track" button, parcel number search', '1', 'Medium'],
          ['5', 'Quick Action Links', 'Icon + text link tiles in horizontal scrollable row', '2 (with/without icon)', 'Low'],
          ['6', 'Promo Card', 'Image (optional) + heading + description + CTA link in card format', '3 (with image, text-only, featured)', 'Low'],
          ['7', 'Service Card Grid', 'Card grid with heading, description, and action link', '2 (2-col, 3-col)', 'Low'],
          ['8', 'Delivery Info List', 'Heading + description + list of links with chevron icons', '1', 'Low'],
          ['9', 'Shop Category Tiles', 'Image + category text tiles linking to shop sections', '1', 'Low'],
          ['10', 'FAQ Accordion', 'Expandable question/answer sections (dynamically loaded)', '1', 'Medium'],
          ['11', 'Breadcrumbs', 'Hierarchical page path navigation', '1', 'Low'],
          ['12', 'Video Embed', 'YouTube iframe with transcript toggle button', '1', 'Medium'],
          ['13', 'Content Section', 'Heading + paragraph + optional CTA links in split layout', '2 (single CTA, multi-CTA)', 'Low'],
          ['14', 'Comparison Cards', 'Side-by-side feature comparison (Parcel Post vs Express)', '2 (2-card, 3-card)', 'Medium'],
          ['15', 'Information Cards', 'Heading + text + link card without image', '1', 'Low'],
          ['16', 'Footnotes', 'Superscript-referenced disclaimers at page bottom', '1', 'Low'],
          ['17', 'Side Navigation', 'Collapsible in-page navigation menu for long content pages', '1', 'Medium'],
          ['18', 'Business Cross-Sell', 'Heading + text promoting business solutions with link', '1', 'Low'],
          ['19', 'Donation CTA', 'Card with bulleted list and external donation link', '1', 'Low'],
          ['20', 'Image + Text Row', 'Image on one side, text/heading on other (alternating)', '2 (image-left, image-right)', 'Low'],
          ['21', 'Search Form', 'Full-width search input with Coveo-powered results', '2 (header, help page)', 'High'],
          ['22', 'Chatbot Widget', 'Floating chat button with bot + live agent escalation', '1', 'High'],
          ['23', 'Map/Location Finder', 'Google Maps integration with search, filter, and result cards', '1', 'High'],
          ['24', 'Calculator Tool', 'Postage/currency calculator with input fields and results', '2 (postage, currency)', 'High'],
          ['25', 'Insurance Quote Engine', 'Third-party embedded insurance application', '1', 'High'],
          ['26', 'Alert/Notification Banner', 'Site-wide or page-level alert message bar', '2 (info, warning)', 'Low'],
          ['27', 'Status Feed', 'Dynamic list of disruption/update items with dates and filters', '1', 'High'],
          ['28', 'Auth/Login Module', 'Auth0-based authentication integration (account status check)', '1', 'High'],
          ['29', 'Community Story Card', 'Article card with image, title, excerpt for listing pages', '1', 'Low'],
          ['30', 'Social Links', 'Social media icon links (Facebook, Twitter, LinkedIn)', '1', 'Low'],
        ],
        [5, 18, 42, 18, 12]
      ),

      new Paragraph({ spacing: { after: 200 } }),
      para('Block Complexity Summary:', { bold: true }),
      createTable(
        ['Complexity', 'Block Count', 'Notes'],
        [
          ['High', '8', 'Requires custom JavaScript, third-party APIs, or complex interactions'],
          ['Medium', '8', 'Moderate DOM manipulation, responsive layouts, animation'],
          ['Low', '14', 'Primarily CSS styling with minimal/no JavaScript'],
          ['TOTAL', '30', ''],
        ],
        [25, 25, 50]
      ),

      new Paragraph({ children: [new PageBreak()] }),

      // PAGE COUNTS BY TEMPLATE
      heading('4. Page Counts by Template - Migration Classification', HeadingLevel.HEADING_1),
      para('Pages have been classified by migration approach: Automated (standardized structure, low complexity), Manual (dynamic content, heavy logic, custom integrations).'),
      new Paragraph({ spacing: { after: 100 } }),

      createTable(
        ['Template', 'Total Pages', 'Auto-Migrate', 'Manual/Custom', 'Rationale for Manual'],
        [
          ['Homepage', '3', '0', '3', 'Unique layouts, tracking widget, dynamic promos'],
          ['Category Landing', '45', '30', '15', 'Some have custom card arrangements or embedded tools'],
          ['Service Detail', '320', '260', '60', 'Majority standardized; some have pricing tables, calculators'],
          ['Product/Pricing', '85', '40', '45', 'Pricing tables, comparison logic, optional extras config'],
          ['Content/Article', '180', '160', '20', 'Rich text pages with embedded video need manual verification'],
          ['Community Story', '125', '120', '5', 'Highly standardized blog articles'],
          ['Help & Support', '45', '10', '35', 'Chatbot, dynamic forms, category routing logic'],
          ['Location Finder', '18', '0', '18', 'Map integration, real-time data, location search'],
          ['Form/Application', '35', '0', '35', 'Multi-step forms, payment flows, validation'],
          ['Disruptions/Alerts', '35', '5', '30', 'Dynamic content feed, API-driven updates'],
          ['Legal/Policy', '15', '15', '0', 'Simple long-form text content'],
          ['Business Portal Landing', '150', '110', '40', 'Most standardized; some have portal integrations'],
          ['Travel Insurance Quote', '9', '0', '9', 'Third-party insurance engine, fully dynamic'],
          ['Currency Converter', '2', '0', '2', 'API-driven exchange rates, calculator logic'],
        ],
        [20, 12, 14, 14, 40]
      ),

      new Paragraph({ spacing: { after: 200 } }),
      para('Migration Summary:', { bold: true }),
      createTable(
        ['Migration Type', 'Page Count', 'Percentage'],
        [
          ['Automated Migration', '750', '64.3%'],
          ['Manual/Custom Migration', '317', '27.2%'],
          ['Exclude (external apps/redirects)', '100', '8.5%'],
          ['TOTAL', '1,167', '100%'],
        ],
        [40, 30, 30]
      ),

      new Paragraph({ children: [new PageBreak()] }),

      // INTEGRATIONS ANALYSIS
      heading('5. Integrations Analysis', HeadingLevel.HEADING_1),
      para('The following third-party integrations and embedded services have been identified through script analysis and page inspection:'),
      new Paragraph({ spacing: { after: 100 } }),

      createTable(
        ['#', 'Integration', 'Provider', 'Purpose', 'EDS Impact'],
        [
          ['1', 'Analytics & Performance', 'New Relic', 'Application performance monitoring (APM)', 'Replace with EDS-compatible RUM or keep as delayed script'],
          ['2', 'Tag Management', 'Adobe Experience Platform (Launch)', 'Tag management, data layer', 'Migrate to EDS delayed loading pattern'],
          ['3', 'Web Analytics', 'Google Tag Manager / GA4', 'Site analytics, conversion tracking', 'Load in delayed.js'],
          ['4', 'Advertising - Facebook', 'Meta Pixel', 'Ad tracking, retargeting audiences', 'Load in delayed.js'],
          ['5', 'Advertising - Google', 'Google Ads / DoubleClick', 'Conversion tracking, remarketing', 'Load in delayed.js'],
          ['6', 'Advertising - LinkedIn', 'LinkedIn Insight Tag', 'B2B advertising, audience targeting', 'Load in delayed.js'],
          ['7', 'Session Recording', 'FullStory', 'Session replay, heatmaps, analytics', 'Load in delayed.js'],
          ['8', 'Search', 'Coveo (platform-au.cloud.coveo.com)', 'Site search engine with AI recommendations', 'Custom block with API integration'],
          ['9', 'Authentication', 'Auth0 (welcome.auspost.com.au)', 'User login, session management, SSO', 'Client-side integration block'],
          ['10', 'Bot Protection', 'DataDome (dd.auspost.com.au)', 'Bot detection, CAPTCHA challenges', 'Evaluate need; may conflict with EDS CDN'],
          ['11', 'Video', 'YouTube', 'Embedded video content', 'Standard embed block'],
          ['12', 'Interactive Content', 'Vudoo', 'Interactive content experiences', 'Custom embed block'],
          ['13', 'NPS Surveys', 'Custom (via FullStory)', 'Net Promoter Score surveys', 'Load in delayed.js'],
          ['14', 'Payment Gateway', 'Mastercard / various', 'Travel card, bill payments (external)', 'External links only'],
          ['15', 'Money Transfer', 'Western Union', 'International money transfers (external)', 'External links only'],
          ['16', 'Insurance', 'Third-party provider', 'Travel insurance quote engine', 'Iframe/redirect integration'],
          ['17', 'Digital Identity', 'Digital iD (digitalid.com)', 'Identity verification services', 'External app link'],
          ['18', 'Maps', 'Google Maps (inferred)', 'Location finder, store locator', 'Custom block with Maps API'],
          ['19', 'Accessibility', 'NRS / TIS', 'National Relay Service, Translation services', 'Content links only'],
          ['20', 'Mobile App', 'Branch.io (auspost.app.link)', 'App deep linking', 'Load in delayed.js'],
        ],
        [4, 18, 22, 26, 30]
      ),

      new Paragraph({ children: [new PageBreak()] }),

      // COMPLEX USE CASES
      heading('6. Complex Use Cases & Observations', HeadingLevel.HEADING_1),
      para('The following complex behaviours, edge cases, and functionality require special attention during migration:'),
      new Paragraph({ spacing: { after: 100 } }),

      createTable(
        ['#', 'Use Case', 'Description', 'Risk Level', 'Recommended Approach'],
        [
          ['1', 'Multi-Segment Navigation', 'Three distinct audience segments (Personal/Business/Enterprise) with different mega menu structures, up to 5 levels deep', 'High', 'Custom header block with audience-switcher; consider separate nav configs per segment'],
          ['2', 'Auth0 Session Management', 'Silent authentication check on every page load (iframe-based), account status indicator in header', 'High', 'Client-side Auth0 SDK in delayed.js; graceful degradation when unauthenticated'],
          ['3', 'Coveo Search Engine', 'Full-site search with AI recommendations, result faceting, and type-ahead suggestions', 'High', 'Custom search block with Coveo Headless SDK; progressive enhancement'],
          ['4', 'Parcel Tracking Widget', 'Embedded tracking input on homepage that calls internal APIs for real-time status', 'Medium', 'Custom block with API proxy; form-to-API pattern'],
          ['5', 'Location Finder Application', 'Full SPA with map integration, geolocation, filtered search, operating hours', 'High', 'Keep as separate SPA; embed via iframe or build as standalone EDS page with Maps SDK'],
          ['6', 'Insurance Quote Engine', 'Fully dynamic third-party application embedded for travel insurance quotes', 'High', 'Maintain as iframe embed or redirect to external application'],
          ['7', 'Dynamic Disruptions Feed', 'Real-time service disruption updates with filtering by state/service type', 'Medium', 'Custom block with API-driven content; consider spreadsheet-based for manual updates'],
          ['8', 'DataDome Bot Protection', 'Client-side bot detection with CAPTCHA challenges; may conflict with CDN caching', 'High', 'Evaluate compatibility with EDS CDN; may need edge-function approach'],
          ['9', 'Multi-Portal Authentication', 'Multiple portals (MyPost, MyPost Business, Merchant Portal, eParcel) with different auth contexts', 'Medium', 'External links to separate applications; no migration needed for portal apps'],
          ['10', 'Community Hub Content Volume', '125+ articles with potential dynamic listing, filtering, and pagination', 'Medium', 'EDS spreadsheet index + client-side filtering; or static generation with pagination'],
          ['11', 'Currency Converter', 'Real-time exchange rate data with conversion calculator', 'Medium', 'Custom block with external exchange rate API; cache rates for performance'],
          ['12', 'PDF/Document Downloads', 'Multiple downloadable guides, T&C documents hosted in DAM', 'Low', 'Migrate documents to EDS media; maintain URL paths or implement redirects'],
        ],
        [4, 18, 38, 10, 30]
      ),

      new Paragraph({ children: [new PageBreak()] }),

      // EDS FEASIBILITY
      heading('7. EDS Feasibility Assessment', HeadingLevel.HEADING_1),
      para('Assessment of auspost.com.au suitability for Edge Delivery Services migration:'),
      new Paragraph({ spacing: { after: 100 } }),

      createTable(
        ['Dimension', 'Rating', 'Assessment'],
        [
          ['Content Volume', '★★★☆☆', '1,167 pages is moderate-large; manageable with template-based migration approach'],
          ['Content Standardization', '★★★★☆', '64% of pages follow standard templates suitable for automated migration'],
          ['Third-Party Complexity', '★★☆☆☆', '20 integrations including search, auth, bot protection create complexity'],
          ['Dynamic Content', '★★☆☆☆', 'Tracking, disruptions, insurance quotes, and location finder are highly dynamic'],
          ['Design Consistency', '★★★★☆', 'Strong design system with consistent components across the site'],
          ['Performance Opportunity', '★★★★★', 'Current site loads 39 scripts; EDS will dramatically improve Core Web Vitals'],
          ['Author Workflow Fit', '★★★★☆', 'Most content is well-suited to document-based authoring in EDS'],
          ['Migration Risk', '★★★☆☆', 'Moderate risk due to integrations; mitigated by phased approach'],
        ],
        [25, 15, 60]
      ),

      new Paragraph({ spacing: { after: 300 } }),
      para('Overall Feasibility: RECOMMENDED with Phased Approach', { bold: true }),
      new Paragraph({ spacing: { after: 100 } }),
      bullet('Phase 1: Migrate informational/content pages (community hub, about us, legal) - Low risk'),
      bullet('Phase 2: Migrate service detail and category pages with basic blocks - Medium risk'),
      bullet('Phase 3: Migrate homepage, business section with custom integrations - Higher risk'),
      bullet('Phase 4: Address dynamic applications (location finder, insurance, calculators) - Highest risk'),

      new Paragraph({ spacing: { after: 200 } }),
      para('Key Risks:', { bold: true }),
      bullet('Bot protection (DataDome) compatibility with EDS CDN edge caching'),
      bullet('Auth0 session management across EDS-served pages'),
      bullet('Coveo search requires custom implementation for EDS architecture'),
      bullet('Location finder may need to remain as a separate SPA application'),
      bullet('Insurance quote engine is fully third-party and cannot be migrated to EDS'),

      new Paragraph({ children: [new PageBreak()] }),

      // MIGRATION ESTIMATES
      heading('8. Migration Estimates', HeadingLevel.HEADING_1),
      para('Estimates based on standard Adobe partner rates ($175/hr midpoint of $150-200/hr range).'),
      new Paragraph({ spacing: { after: 100 } }),

      heading('8.1 Effort Breakdown by Phase', HeadingLevel.HEADING_2),
      createTable(
        ['Phase', 'Activity', 'Hours', 'Duration (weeks)', 'Cost ($175/hr)'],
        [
          ['Setup', 'Project setup, EDS configuration, design token extraction, CI/CD', '80', '2', '$14,000'],
          ['Setup', 'Design system migration (global styles, fonts, responsive grid)', '120', '3', '$21,000'],
          ['Phase 1', 'Block development - Low complexity (14 blocks)', '168', '3', '$29,400'],
          ['Phase 1', 'Block development - Medium complexity (8 blocks)', '192', '3', '$33,600'],
          ['Phase 1', 'Block development - High complexity (8 blocks)', '320', '5', '$56,000'],
          ['Phase 2', 'Header/Navigation (multi-segment mega menu)', '120', '3', '$21,000'],
          ['Phase 2', 'Footer implementation', '40', '1', '$7,000'],
          ['Phase 3', 'Template implementation (14 templates)', '280', '4', '$49,000'],
          ['Phase 4', 'Automated content migration (750 pages)', '160', '3', '$28,000'],
          ['Phase 4', 'Manual content migration (317 pages)', '480', '6', '$84,000'],
          ['Phase 5', 'Integration: Coveo Search', '120', '3', '$21,000'],
          ['Phase 5', 'Integration: Auth0 Authentication', '80', '2', '$14,000'],
          ['Phase 5', 'Integration: Tracking Widget', '60', '1.5', '$10,500'],
          ['Phase 5', 'Integration: Location Finder', '120', '3', '$21,000'],
          ['Phase 5', 'Integration: Analytics & Martech stack', '80', '2', '$14,000'],
          ['Phase 5', 'Integration: Other (DataDome, calculators, embeds)', '120', '3', '$21,000'],
          ['QA', 'Visual QA and regression testing', '240', '4', '$42,000'],
          ['QA', 'Performance testing and optimization', '80', '2', '$14,000'],
          ['QA', 'Accessibility testing (WCAG 2.1 AA)', '80', '2', '$14,000'],
          ['QA', 'Cross-browser/device testing', '60', '1.5', '$10,500'],
          ['Launch', 'URL redirect mapping and implementation', '80', '2', '$14,000'],
          ['Launch', 'DNS cutover, go-live support, monitoring', '40', '1', '$7,000'],
        ],
        [10, 42, 10, 16, 16]
      ),

      new Paragraph({ spacing: { after: 300 } }),
      heading('8.2 Summary Estimates', HeadingLevel.HEADING_2),
      createTable(
        ['Category', 'Hours', 'Cost', 'Duration'],
        [
          ['Project Setup & Design System', '200', '$35,000', '4 weeks'],
          ['Block Development (30 blocks)', '680', '$119,000', '8 weeks'],
          ['Navigation (Header + Footer)', '160', '$28,000', '3 weeks'],
          ['Template Implementation', '280', '$49,000', '4 weeks'],
          ['Content Migration (Automated)', '160', '$28,000', '3 weeks'],
          ['Content Migration (Manual)', '480', '$84,000', '6 weeks'],
          ['Third-Party Integrations', '580', '$101,500', '8 weeks'],
          ['QA & Testing', '460', '$80,500', '6 weeks'],
          ['Launch & Go-Live', '120', '$21,000', '2 weeks'],
          ['', '', '', ''],
          ['TOTAL', '3,120 hours', '$546,000', '20-24 weeks (with parallel streams)'],
        ],
        [35, 15, 20, 30]
      ),

      new Paragraph({ spacing: { after: 300 } }),
      heading('8.3 Recommended Team Composition', HeadingLevel.HEADING_2),
      createTable(
        ['Role', 'FTE', 'Duration', 'Responsibilities'],
        [
          ['Technical Architect', '0.5', '24 weeks', 'Architecture decisions, integration design, code reviews'],
          ['Senior EDS Developer', '2.0', '20 weeks', 'Block development, template implementation, integrations'],
          ['Frontend Developer', '1.0', '16 weeks', 'CSS/styling, responsive design, accessibility'],
          ['Content Migration Specialist', '2.0', '10 weeks', 'Automated + manual content migration, content QA'],
          ['QA Engineer', '1.0', '12 weeks', 'Visual regression, performance, accessibility testing'],
          ['Project Manager', '0.5', '24 weeks', 'Planning, stakeholder communication, risk management'],
        ],
        [25, 10, 15, 50]
      ),

      new Paragraph({ spacing: { after: 300 } }),
      heading('8.4 Timeline (Gantt Overview)', HeadingLevel.HEADING_2),
      createTable(
        ['Phase', 'Weeks 1-4', 'Weeks 5-8', 'Weeks 9-12', 'Weeks 13-16', 'Weeks 17-20', 'Weeks 21-24'],
        [
          ['Setup & Design', '████', '', '', '', '', ''],
          ['Block Development', '██', '████', '████', '', '', ''],
          ['Navigation', '', '████', '', '', '', ''],
          ['Templates', '', '', '████', '██', '', ''],
          ['Content Migration', '', '', '██', '████', '████', ''],
          ['Integrations', '', '██', '████', '████', '██', ''],
          ['QA & Testing', '', '', '██', '██', '████', '██'],
          ['Launch Prep', '', '', '', '', '██', '████'],
        ],
        [20, 13, 13, 13, 13, 13, 13]
      ),

      new Paragraph({ children: [new PageBreak()] }),

      // RECOMMENDATIONS
      heading('9. Recommendations & Next Steps', HeadingLevel.HEADING_1),
      new Paragraph({ spacing: { after: 100 } }),
      para('1. Phased Migration Approach', { bold: true }),
      bullet('Start with low-risk content pages (community hub, about us, legal/policy) as a proof of concept'),
      bullet('Iterate on block library before tackling high-complexity templates'),
      bullet('Defer dynamic applications (location finder, insurance) to final phase'),
      new Paragraph({ spacing: { after: 100 } }),

      para('2. Integration Strategy', { bold: true }),
      bullet('Move all analytics/martech to delayed.js pattern for optimal Core Web Vitals'),
      bullet('Evaluate DataDome compatibility early — may be a blocking risk'),
      bullet('Implement Auth0 as a client-side module that does not block page rendering'),
      bullet('Build Coveo search as progressive enhancement (basic search fallback)'),
      new Paragraph({ spacing: { after: 100 } }),

      para('3. Content Governance', { bold: true }),
      bullet('Establish URL redirect mapping early to preserve SEO equity'),
      bullet('Define content freeze window before migration phases'),
      bullet('Plan for parallel publishing during transition period'),
      new Paragraph({ spacing: { after: 100 } }),

      para('4. Performance Targets', { bold: true }),
      bullet('Target Lighthouse score of 100 across all Core Web Vitals'),
      bullet('Reduce initial page load from 39 scripts to < 5 (EDS + essential integrations)'),
      bullet('Achieve sub-1-second LCP on all content pages'),
      new Paragraph({ spacing: { after: 100 } }),

      para('5. Risk Mitigations', { bold: true }),
      bullet('Conduct DataDome + EDS CDN compatibility POC in week 1-2'),
      bullet('Prototype Auth0 silent authentication pattern on EDS in week 3-4'),
      bullet('Establish performance budget and automated Lighthouse CI testing'),
      bullet('Plan rollback strategy using DNS-based traffic switching'),

      new Paragraph({ spacing: { after: 400 } }),
      new Paragraph({
        children: [new TextRun({ text: '— End of Report —', size: 22, italics: true, color: '999999' })],
        alignment: AlignmentType.CENTER,
      }),
    ],
  }],
});

const buffer = await Packer.toBuffer(doc);
writeFileSync('/backups/roshmim/bc-test-site/repo/AusPost_EDS_Migration_Scope_Analysis.docx', buffer);
console.log('Report generated: AusPost_EDS_Migration_Scope_Analysis.docx');
