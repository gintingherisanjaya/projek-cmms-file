/**
 * Shared configuration for Func Loc scripts (template applier + runner).
 *
 * Central place to keep Drive folder URL and allowed owner emails in sync.
 */

/** NEW_REGISTRASI NAMA MESIN ALAT PKS – Drive source/target folder */
export const NEW_REGISTRASI_DRIVE_FOLDER_URL =
  'https://drive.google.com/drive/folders/1jlNypjanMeK8Q3ssW0P_QnfoKRMoS9K7';
  // 'https://drive.google.com/drive/folders/1cCSNlTObRnmUz_7NOO7saYtYh2ehVjys';

/**
 * Allowed owners for PKS files in the Drive folder.
 * Union of files owned by any of these is used as the working set.
 */
export const SOURCE_DRIVE_OWNER_EMAILS = ['pkskonsultan@gmail.com', 'butarbutar150303@gmail.com'];
// export const SOURCE_DRIVE_OWNER_EMAILS = ['fahrurrozy4214@gmail.com'];

/** BAH JAMBI / master data reference (Google Sheets) – template source */
export const BAH_JAMBI_SHEETS_URL =
  'https://docs.google.com/spreadsheets/d/1yjet6YmtkY3r1ITuw7DTRuMebEAa5maq/edit?usp=sharing';

/** Before/after DESC+EQKTU alias mapping (Google Sheets) */
export const BEFORE_AFTER_DESC_FIX_SHEETS_URL =
  'https://docs.google.com/spreadsheets/d/1usGu0VzXyHnHWQe9ARR_noVnGR5C-_YdBOlCpTc_To4/edit?gid=0#gid=0';


