# AEM Edge Delivery Services Migration Analysis Plan — auspost.com.au

## Objective

Perform a comprehensive site analysis of **auspost.com.au** through live crawling to produce a detailed migration assessment document (MS Word format) covering templates, blocks, page counts, integrations, complex use cases, and effort estimates using standard Adobe partner rates (~$150-200/hr).

## Approach

### Phase 1: URL Discovery & Site Crawl
- Crawl the sitemap(s) at auspost.com.au to discover all accessible URLs
- Categorize URLs by path patterns to identify template groupings
- Capture representative pages from each template category

### Phase 2: Template Identification
- Analyze page structures to identify unique templates
- Classify templates by complexity (Low/Medium/High)
- Count pages per template type

### Phase 3: Block/Component Cataloging
- Identify all reusable blocks and components across sampled pages
- Distinguish design variations (same content model, different visual) from distinct blocks
- Classify each block by complexity

### Phase 4: Integration & Complexity Analysis
- Identify third-party integrations (analytics, forms, payment, auth, APIs)
- Flag complex use cases requiring special migration attention
- Assess EDS feasibility for each area

### Phase 5: Estimation & Document Generation
- Calculate automated vs. manual migration effort
- Produce timeline and cost estimates at standard Adobe rates
- Generate formatted MS Word document with all findings

## Deliverables

| Deliverable | Format |
|---|---|
| Templates Inventory Table | Word table |
| Blocks/Components Catalog | Word table |
| Page Counts by Template | Word table |
| Integrations Analysis | Word table |
| Complex Use Cases & Observations | Word section |
| EDS Feasibility Comparison | Word table |
| Migration Estimates (hours, cost, timeline) | Word table |

## Assumptions

- Standard Adobe partner rates: $150–200/hr (will use $175/hr midpoint)
- Live crawl limited to publicly accessible pages (no authenticated content)
- Analysis covers desktop and mobile responsive layouts
- Dynamic/application pages (e.g., tracking, account login) identified but not deeply analyzed for content migration

## Checklist

- [ ] Discover URLs via sitemap crawl of auspost.com.au
- [ ] Categorize URLs by path pattern to identify template groups
- [ ] Analyze representative pages from each template group
- [ ] Identify and catalog all unique page templates with complexity ratings
- [ ] Identify and catalog all reusable blocks/components with complexity ratings
- [ ] Count pages per template and classify auto vs. manual migration
- [ ] Identify third-party integrations and embedded services
- [ ] Document complex use cases and edge cases
- [ ] Assess EDS feasibility, risks, and complexity
- [ ] Calculate migration effort estimates (automated, manual, QA)
- [ ] Produce cost and timeline estimates at standard Adobe rates
- [ ] Generate MS Word document with all tables and analysis

## Execution Requirement

This plan requires **Execute mode** to perform live crawls, page analysis, and document generation. Switch to Execute mode to begin implementation.
