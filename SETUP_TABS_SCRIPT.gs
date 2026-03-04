function setupMPATabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup _prompts tab
  var prompts = ss.getSheetByName('_prompts');
  if (!prompts) prompts = ss.insertSheet('_prompts');
  prompts.clear();
  prompts.getRange(1, 1, 1, 7).setValues([['prompt_id', 'stage', 'sub_analysis_type', 'prompt_text', 'version', 'last_updated', 'active']]);
  prompts.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
  var promptData = [
    ['P1', '2', 'general_web_search', 'You are a market research analyst at Zuora. Given this product idea: {idea}. Search for market trends, existing solutions, and opportunities in the subscription/commerce space. Provide key_insights (array), opportunities (array), and risks (array). Return JSON.', '1.0', '2026-03-04', 'TRUE'],
    ['P2', '2', 'competitor_eval', 'You are a competitive intelligence analyst. Compare this idea against these competitors: {competitor_list}. For each competitor assess: feature_available (yes/no/partial), zuora_status (available/planned/gap), gap_level (none/minor/major). Return JSON with competitive_matrix array.', '1.0', '2026-03-04', 'TRUE'],
    ['P3', '2', 'adjacent_solutions', 'You are an innovation scout. For this problem: {idea}. Identify companies in OTHER industries solving similar problems. Focus on transferable mechanisms and inspiration. Return JSON with adjacent_solutions array of {industry, company, solution, transferable_idea}.', '1.0', '2026-03-04', 'TRUE'],
    ['P4', '2', 'customer_rag', 'You are a customer insights analyst at Zuora. Given this idea: {idea}. Analyze the following customer evidence to identify pain points, feature requests, severity levels, and customer segments affected. Return JSON with customer_evidence array and demand_signals.', '1.0', '2026-03-04', 'TRUE'],
    ['P5', '2', 'tam_snapshot', 'You are a market sizing analyst. For this product idea: {idea}. Estimate TAM, SAM, and SOM for the addressable market. Include growth rates (CAGR), key segments, and rationale. Return JSON with tam_usd, sam_usd, som_usd, cagr, key_markets array.', '1.0', '2026-03-04', 'TRUE']
  ];
  prompts.getRange(2, 1, promptData.length, 7).setValues(promptData);
  prompts.autoResizeColumns(1, 7);

  // 2. Create _projects tab
  var projects = ss.getSheetByName('_projects');
  if (!projects) projects = ss.insertSheet('_projects');
  projects.clear();
  projects.getRange(1, 1, 1, 11).setValues([['project_id', 'name', 'owner_email', 'description', 'google_folder_id', 'google_form_id', 'form_url', 'stage', 'feasibility_score', 'status', 'created_date']]);
  projects.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
  projects.getRange(2, 1, 1, 11).setValues([['PROJ-001', 'Merchandising V1', 'diversifiedbusinessmgmt@gmail.com', 'Zuora Commerce Merchandising capabilities', '', '', '', 'research', '7.87', 'active', '2026-02-28']]);
  projects.autoResizeColumns(1, 11);

  // 3. Create _stage2_results tab
  var stage2 = ss.getSheetByName('_stage2_results');
  if (!stage2) stage2 = ss.insertSheet('_stage2_results');
  stage2.clear();
  stage2.getRange(1, 1, 1, 6).setValues([['project_id', 'analysis_type', 'results_json', 'score', 'timestamp', 'status']]);
  stage2.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
  stage2.autoResizeColumns(1, 6);

  // 4. Create _competitor_registry tab
  var competitors = ss.getSheetByName('_competitor_registry');
  if (!competitors) competitors = ss.insertSheet('_competitor_registry');
  competitors.clear();
  competitors.getRange(1, 1, 1, 7).setValues([['competitor_id', 'name', 'website_url', 'doc_count_in_rag', 'last_scraped', 'status', 'active']]);
  competitors.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
  var compData = [
    ['COMP-001', 'Salesforce', 'https://www.salesforce.com/products/commerce/', '0', '', 'pending', 'TRUE'],
    ['COMP-002', 'BillingPlatform', 'https://billingplatform.com/', '0', '', 'pending', 'TRUE'],
    ['COMP-003', 'Metronome', 'https://metronome.com/', '0', '', 'pending', 'TRUE'],
    ['COMP-004', 'Orb', 'https://www.withorb.com/', '0', '', 'pending', 'TRUE'],
    ['COMP-005', 'Stripe Billing', 'https://stripe.com/billing', '0', '', 'pending', 'TRUE']
  ];
  competitors.getRange(2, 1, compData.length, 7).setValues(compData);
  competitors.autoResizeColumns(1, 7);

  // 5. Create _segmentation_tags tab
  var tags = ss.getSheetByName('_segmentation_tags');
  if (!tags) tags = ss.insertSheet('_segmentation_tags');
  tags.clear();
  tags.getRange(1, 1, 1, 4).setValues([['tag_id', 'category', 'label', 'description']]);
  tags.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
  var tagData = [
    ['SEG-001', 'business_model', 'B2B', 'Business to Business'],
    ['SEG-002', 'business_model', 'B2C', 'Business to Consumer'],
    ['SEG-003', 'business_model', 'B2B2C', 'Business to Business to Consumer'],
    ['SEG-004', 'company_size', 'Enterprise', 'Enterprise (1000+ employees)'],
    ['SEG-005', 'company_size', 'Mid-Market', 'Mid-Market (100-999 employees)'],
    ['SEG-006', 'company_size', 'SMB', 'Small/Medium Business (1-99 employees)'],
    ['SEG-007', 'relationship', 'Existing Customer', 'Current Zuora customer'],
    ['SEG-008', 'relationship', 'Prospect', 'Potential new customer'],
    ['SEG-009', 'relationship', 'Churned', 'Former customer']
  ];
  tags.getRange(2, 1, tagData.length, 4).setValues(tagData);
  tags.autoResizeColumns(1, 4);

  SpreadsheetApp.flush();
  Logger.log('All 5 MPA tabs created and seeded successfully!');
}
