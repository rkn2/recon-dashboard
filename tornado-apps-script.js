// Google Apps Script — Tornado Damage DB backend.
// Paste this into your Google Sheet's script editor.
//
// SETUP:
// 1. Create a new Google Sheet (this will be the tornado damage database)
// 2. Extensions → Apps Script
// 3. Paste this entire file, replacing the default code
// 4. Click Deploy → New deployment
//    - Type: Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 5. Copy the deployment URL (looks like https://script.google.com/macros/s/.../exec)
// 6. Paste it into SYNC_ENDPOINT in tornado.html (search for SYNC_ENDPOINT)
//
// The sheet auto-creates a "Buildings" tab with all column headers on first
// request. Each building is one row, keyed by building_id (column A).
// Saves upsert: existing buildings update in place, new ones append.
//
// SECURITY NOTE: "Anyone" means anyone with the URL can read/write building
// records. This is acceptable for the student data-entry workflow on a
// public GitHub Pages site, but don't put sensitive data in the sheet.

var SHEET_NAME = 'Buildings';

var HEADERS = [
  'building_id',
  'completed_by',
  'last_modified',
  'S.No.',
  'ref# (DELETE LATER)',
  'tornado_name',
  'tornado_year',
  'town',
  'overall_photos_link',
  'tornado_EF',
  'tornado_start_lat',
  'tornado_end_lat',
  'tornado_start_long',
  'tornado_end_long',
  'latitude',
  'longitude',
  'complete_address',
  'building_name_listing',
  'building_name_current',
  'archetype',
  'located_in_historic_district',
  'occupany_u',
  'number_stories',
  'year_built_u',
  'building_area_m2',
  'building_urban_setting',
  'building_position_on_street',
  'buidling_height_m',
  'first_floor_elevation_m',
  'front_elevation_orientation',
  'wall_length_side',
  'wall_length_front',
  'wall_thickness',
  'roof_shape_u',
  'roof_slope_u',
  'construction_type_u',
  'construction_type_u_unc',
  'mwfrs_u_wall',
  'mwfrs_u_wall_unc',
  'mwfrs_u_roof',
  'mwfrs_u_roof_unc',
  'mwfrs_u_moment_frame',
  'mwfrs_u_unc',
  'masonry_leaves',
  'masonry_leaves_unc',
  'structural_wall_system_u',
  'foundation_type_u',
  'foundation_type_u_unc',
  'wall_anchorage_type_u',
  'wall_anchorage_type_u_unc',
  'wall_substrate_u',
  'wall_substrate_u_unc',
  'wall_cladding_u',
  'wall_cladding_u_unc',
  'soffit_type_u',
  'soffit_present_u',
  'wall_fenesteration_per_front',
  'wall_fenesteration_per_back',
  'wall_fenesteration_per_right',
  'wall_fenesteration_per_left',
  'wall_fenestration_per_n',
  'wall_fenestration_per_s',
  'wall_fenestration_per_e',
  'wall_fenestration_per_w',
  'wall_fenesteration_protection_front',
  'wall_fenesteration_protection_front_unc',
  'wall_fenesteration_protection_back',
  'wall_fenesteration_protection_back_unc',
  'wall_fenesteration_protection_right',
  'wall_fenesteration_protection_right_unc',
  'wall_fenesteration_protection_left',
  'wall_fenesteration_protection_left_unc',
  'fenestration_protection_type_front',
  'fenestration_protection_type_back',
  'fenestration_protection_type_right',
  'fenestration_protection_type_left',
  'fenestration_protection_type_unc',
  'large_door_present_front',
  'large_door_present_back',
  'large_door_present_right',
  'large_door_present_left',
  'door_present_n',
  'door_present_s',
  'door_present_e',
  'door_present_w',
  'large_door_opening_type_',
  'opening_type',
  'roof_system_u',
  'roof_system_u_unc',
  'r2wall_attachment_u',
  'r2wall_attachment_type_unc',
  'roof_substrate_type_u',
  'roof_substrate_type_u_unc',
  'roof_cover_u',
  'overhang_length_u',
  'parapet_height_m',
  'secondary_water_barrier_u',
  'retrofit_present_u',
  'retrofit_year_u',
  'retrofit_type_u',
  'retrofit_type_unc_u',
  'retrofit_type_1_u',
  'reotrofit_2_u',
  'hazards_present_u',
  'status_u',
  'wind_damage_rating_u',
  'damage_indicator_u',
  'degree_of_damage_u',
  'wind_damage_details_u',
  'roof_structure_damage_u',
  'roof_structure_damage_u_per',
  'roof_susbtrate_damage_u',
  'roof_substrate_damage_per',
  'foundation_failure_u',
  'foundatio_failure_per_u',
  'wall_structure_damage_u',
  'wall_structure_damage_per_front',
  'wall_structure_damage_per_back',
  'wall_structure_damage_per_left',
  'wall_structure_damage_per_right',
  'wall_structure_damage_n',
  'wall_structure_damage_s',
  'wall_structure_damage_e',
  'wall_structure_damage_w',
  'wall_substrate_damage_u',
  'wall_substrate_damage_per_front',
  'wall_substrate_damage_per_back',
  'wall_substrate_damage_per_right',
  'wall_substrate_damage_per_left',
  'wall_substrate_damage_n',
  'wall_substrate_damage_s',
  'wall_substrate_damage_e',
  'wall_substrate_damage_w',
  'wall_cladding_damage_per_front',
  'wall_cladding_damage_per_back',
  'wall_cladding_damage_per_right',
  'wall_cladding_damage_per_left',
  'wall_cladding_damage_n',
  'wall_cladding_damage_s',
  'wall_cladding_damage_e',
  'wall_cladding_damage_w',
  'damaged_fenesteration_per_front',
  'damaged_fenesteration_per_back',
  'damaged_fenesteration_per_right',
  'damaged_fenesteration_per_left',
  'wall_fenestration_damage_per_n',
  'wall_fenestration_damage_per_s',
  'wall_fenestration_damage_per_e',
  'wall_fenestration_damage_per_w',
  'soffit_damage_per_u',
  'fascia_damage_per_u',
  'piles_damage_u',
  'foundation_damage_cause_u',
  'foundation_damage_u',
  'foundation_damage_u_per',
  'oop_failure_n',
  'oop_failure_s',
  'oop_failure_e',
  'oop_failure_w',
  'ip_failure_n',
  'ip_failure_s',
  'ip_failure_e',
  'ip_failure_w',
  'surge_damage_rating_u',
  'rainwater_ingress_damage_rating_u',
  'wall_cladding_damage_level_u',
  ' stories_with_damage_u',
  'per_building_footprint_eroded_u',
  'damage_status',
  'overall_photos_front',
  'overall_photos_back',
  'overall_photos_left',
  'overall_photos_right',
  'detailed_photos_front',
  'detailed_photos_back',
  'detailed_photos_left',
  'detailed_photos_right',
  'national_register_listing_year',
  'existed_during_tornado',
  'buidling_use_before_tornado',
  'buidling_use_after_tornado',
  'buidling_use_plan_after_tornado',
  'demolishing_year',
  'single_unit',
  'multiple_unit',
  'world_heritage_property',
  'hague_convention',
  'sub_national_heritage _list',
  'iucn_protected_area',
  'property_of_local_significance',
  'buidling_existed_5_yrs_before_tornado',
  'buidling_existed_3_yrs_before_tornado',
  'buidling_existed_1_yrs_before_tornado',
  'building_existed_during_tornado',
  'building_in_use_during_tornado',
  'building_use_during_tornado',
  'building_demolished_1_yrs_after_tornado',
  'building_demolished_3_yrs_after_tornado',
  'building_demolished_5_yrs_after_tornado',
  'const_material_h_stone',
  'const_material_h_brick',
  'const_material_h_wood',
  'const_material_h_mud',
  'const_material_h_rf_masonry',
  'const_material_h_rglr_stone',
  'const_material_h_rf_conc',
  'const_material_h_ir_stone',
  'const_material_h_othr',
  'const_material_v_stone',
  'const_material_v_brick',
  'const_material_v_wood',
  'const_material_v_mud',
  'const_material_v_rf_masonry',
  'const_material_v_rglr_stone',
  'const_material_v_rf_conc',
  'const_material_v_ir_stone',
  'const_material_v_othr',
  'prop_agricultural ',
  'prop_cave',
  'prop_culture _entertainment_facility',
  'prop_forest',
  'prop_industrial_facility',
  'prop_lake',
  'prop_military',
  'prop_nature',
  'prop_religious',
  'prop_rock formation',
  'prop_sports_facility',
  'prop_unilities_facility',
  'prop_archaeological',
  'prop_commemorative structure or landmark',
  'prop_ecosystem',
  'prop_habitat',
  'prop_infrastructure',
  'prop_law / government facility',
  'prop_mine',
  'prop_park / garden',
  'prop_residential facility',
  'prop_scenic area',
  'prop_transportation facility',
  'prop_volcano',
  'prop_battlefield',
  'prop_commercial / exchange facility',
  'prop_educational facility',
  'prop_health / welfare facility',
  'prop_island(s)',
  'prop_marine zone',
  'prop_mountain',
  'parking / storage facility',
  'river catchment system',
  'sea',
  'underground facility',
  'zoological park',
  'prop_val_evidential',
  'prop_val_historical',
  'prop_val_aesthetic',
  'prop_val_communal',
  'owner_individual',
  'owner_business',
  'owner_government',
  'owner_ngo',
  'owner_religious ',
  'owner_unknown',
  'risk_category_16',
  'building_low_rise',
  'building_open',
  'building_enclosed',
  'building_regular_shape',
  'exposure_category',
  'building_rigid',
  'understory_pct_of_building_footprint',
  'notes',
];

// ── Helpers ─────────────────────────────────────────────────────────

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function headerIndex(sheet) {
  var row = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < row.length; i++) {
    map[row[i]] = i;
  }
  return map;
}

function rowToObj(headers, row) {
  var obj = {};
  for (var i = 0; i < headers.length; i++) {
    var val = i < row.length ? row[i] : '';
    if (val !== '') obj[headers[i]] = val;
  }
  return obj;
}

function findRowById(sheet, id) {
  if (!id && id !== 0) return -1;
  var idStr = String(id);
  var data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === idStr) return i + 2;
  }
  return -1;
}

function nextId(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < ids.length; i++) {
    var n = Number(ids[i][0]);
    if (n > max) max = n;
  }
  return max + 1;
}

// ── GET: read buildings ─────────────────────────────────────────────

function doGet(e) {
  try {
    var sheet = getOrCreateSheet();
    var params = e ? e.parameter : {};

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      if (params.id) return jsonOut({ status: 'ok', building: null });
      return jsonOut({ status: 'ok', buildings: [] });
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
    var rows = dataRange.getValues();

    if (params.id) {
      var idStr = String(params.id);
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]) === idStr) {
          return jsonOut({ status: 'ok', building: rowToObj(headers, rows[i]) });
        }
      }
      return jsonOut({ status: 'ok', building: null });
    }

    var buildings = [];
    for (var j = 0; j < rows.length; j++) {
      buildings.push(rowToObj(headers, rows[j]));
    }
    return jsonOut({ status: 'ok', count: buildings.length, buildings: buildings });

  } catch (err) {
    return jsonOut({ status: 'error', message: err.message });
  }
}

// ── POST: upsert a building ─────────────────────────────────────────

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return jsonOut({ status: 'error', message: 'Server busy, try again in a moment.' });
  }

  try {
    var payload = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet();
    var hdrMap = headerIndex(sheet);

    var id = payload.building_id;
    var rowNum = id ? findRowById(sheet, id) : -1;
    var isNew = (rowNum === -1);

    if (isNew) {
      id = nextId(sheet);
      payload.building_id = id;
    }

    payload.last_modified = new Date().toISOString();

    var rowData = [];
    for (var i = 0; i < HEADERS.length; i++) {
      var key = HEADERS[i];
      if (key in payload) {
        rowData.push(payload[key]);
      } else if (!isNew) {
        rowData.push(null);
      } else {
        rowData.push('');
      }
    }

    if (isNew) {
      sheet.appendRow(rowData);
    } else {
      var existingRow = sheet.getRange(rowNum, 1, 1, HEADERS.length).getValues()[0];
      for (var j = 0; j < HEADERS.length; j++) {
        if (rowData[j] === null) {
          rowData[j] = existingRow[j];
        }
      }
      sheet.getRange(rowNum, 1, 1, HEADERS.length).setValues([rowData]);
    }

    lock.releaseLock();
    return jsonOut({
      status: 'ok',
      action: isNew ? 'created' : 'updated',
      building_id: id
    });

  } catch (err) {
    lock.releaseLock();
    return jsonOut({ status: 'error', message: err.message });
  }
}
