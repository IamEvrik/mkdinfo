-- -*- coding: cp1251 -*-
DO $$
BEGIN
	IF NOT EXISTS (SELECT 1 FROM pg_type WHERE typname = 'access_type') THEN
	   CREATE TYPE ACCESS_TYPE AS ENUM ('gwt');
	END IF;
END $$;

DROP FUNCTION IF EXISTS get_municipal_district;
DROP FUNCTION IF EXISTS get_municipal_districts;
DROP FUNCTION IF EXISTS create_municipal_district;
DROP FUNCTION IF EXISTS change_municipal_district;
DROP FUNCTION IF EXISTS delete_municipal_district;
DROP FUNCTION IF EXISTS get_villages;
DROP FUNCTION IF EXISTS get_village;
DROP FUNCTION IF EXISTS change_village;
DROP FUNCTION IF EXISTS delete_village;
DROP FUNCTION IF EXISTS create_village;
DROP FUNCTION IF EXISTS getStreet;
DROP FUNCTION IF EXISTS changeStreet;
DROP FUNCTION IF EXISTS createStreet;
DROP FUNCTION IF EXISTS deleteStreet;
DROP FUNCTION IF EXISTS getDogovor;
DROP FUNCTION IF EXISTS changeDogovor;
DROP FUNCTION IF EXISTS createDogovor;
DROP FUNCTION IF EXISTS deleteDogovor;
DROP FUNCTION IF EXISTS get_contractors;
DROP FUNCTION IF EXISTS get_contractor;
DROP FUNCTION IF EXISTS change_contractor;
DROP FUNCTION IF EXISTS create_contractor;
DROP FUNCTION IF EXISTS delete_contractor;
DROP FUNCTION IF EXISTS getGWT;
DROP FUNCTION IF EXISTS changeGWT;
DROP FUNCTION IF EXISTS createGWT;
DROP FUNCTION IF EXISTS deleteGWT;
DROP FUNCTION IF EXISTS getImprovement;
DROP FUNCTION IF EXISTS changeImprovement;
DROP FUNCTION IF EXISTS createImprovement;
DROP FUNCTION IF EXISTS deleteImprovement;
DROP FUNCTION IF EXISTS getPlanStatus;
DROP FUNCTION IF EXISTS getWallMaterial;
DROP FUNCTION IF EXISTS changeWallMaterial;
DROP FUNCTION IF EXISTS createWallMaterial;
DROP FUNCTION IF EXISTS deleteWallMaterial;
DROP FUNCTION IF EXISTS getWorkType;
DROP FUNCTION IF EXISTS changeWorkType;
DROP FUNCTION IF EXISTS createWorkType;
DROP FUNCTION IF EXISTS deleteWorkType;
DROP FUNCTION IF EXISTS getWorkKind;
DROP FUNCTION IF EXISTS changeWorkKind;
DROP FUNCTION IF EXISTS createWorkKind;
DROP FUNCTION IF EXISTS deleteWorkKind;
DROP FUNCTION IF EXISTS getWorkKindsByWT;
DROP FUNCTION IF EXISTS getBldnIdNoList;
DROP FUNCTION IF EXISTS get_mc;
DROP FUNCTION IF EXISTS changeMC;
DROP FUNCTION IF EXISTS createMC;
DROP FUNCTION IF EXISTS deleteMC;
DROP FUNCTION IF EXISTS getEmployee;
DROP FUNCTION IF EXISTS changeEmployee;
DROP FUNCTION IF EXISTS createEmployee;
DROP FUNCTION IF EXISTS deleteEmployee;
DROP FUNCTION IF EXISTS getEmployeesInOrganization;
DROP FUNCTION IF EXISTS get_plan_statuses;
DROP FUNCTION IF EXISTS get_plan_statuses_new_work;
DROP FUNCTION IF EXISTS get_plan_work;
DROP FUNCTION IF EXISTS change_plan_work;
DROP FUNCTION IF EXISTS create_plan_work;
DROP FUNCTION IF EXISTS delete_plan_work;
DROP FUNCTION IF EXISTS get_bldn_plan_years;
DROP FUNCTION IF EXISTS get_plan_works_by_bldn;
DROP FUNCTION IF EXISTS get_bldn_types;
DROP FUNCTION IF EXISTS get_building;
DROP FUNCTION IF EXISTS create_building;
DROP FUNCTION IF EXISTS change_bldn_services;
DROP FUNCTION IF EXISTS change_bldn_common;
DROP FUNCTION IF EXISTS change_bldn_dogovor;
DROP FUNCTION IF EXISTS delete_building;
DROP FUNCTION IF EXISTS get_bldn_tech_info;
DROP FUNCTION IF EXISTS change_bldn_tech_info;
DROP FUNCTION IF EXISTS getBuildingLandInfo;
DROP FUNCTION IF EXISTS change_bldn_land_info;
DROP FUNCTION IF EXISTS get_work;
DROP FUNCTION IF EXISTS create_work;
DROP FUNCTION IF EXISTS change_work;
DROP FUNCTION IF EXISTS delete_work;
DROP FUNCTION IF EXISTS open_next_term;
DROP FUNCTION IF EXISTS get_fsource;
DROP FUNCTION IF EXISTS get_fsources;
DROP FUNCTION IF EXISTS change_fsource;
DROP FUNCTION IF EXISTS create_fsource;
DROP FUNCTION IF EXISTS delete_fsource;
DROP FUNCTION IF EXISTS get_bldn_works;
DROP FUNCTION IF EXISTS getBldnWorkYears;
DROP FUNCTION IF EXISTS getWorkYears;
DROP FUNCTION IF EXISTS create_sheet;
DROP FUNCTION IF EXISTS get_avr_period;
DROP FUNCTION IF EXISTS load_avr;
DROP FUNCTION IF EXISTS get_expense_items;
DROP FUNCTION IF EXISTS get_expense_item;
DROP FUNCTION IF EXISTS change_expense_item;
DROP FUNCTION IF EXISTS create_expense_item;
DROP FUNCTION IF EXISTS delete_expense_item;
DROP FUNCTION IF EXISTS change_expense;
DROP FUNCTION IF EXISTS delete_expense;
DROP FUNCTION IF EXISTS add_expense;
DROP FUNCTION IF EXISTS copy_expenses_from_term;
DROP FUNCTION IF EXISTS delete_expenses_in_term;
DROP FUNCTION IF EXISTS delete_bldn_expenses;
DROP FUNCTION IF EXISTS bldn_list_use_expense_name;
DROP FUNCTION IF EXISTS bldn_change_expense_name;
DROP FUNCTION IF EXISTS bldn_delete_expense_name;
DROP FUNCTION IF EXISTS	bldn_add_expense_name;
DROP FUNCTION IF EXISTS get_bldn_last_expenses;
DROP FUNCTION IF EXISTS get_bldn_expense_history;
DROP FUNCTION IF EXISTS get_bldn_expenses_in_term;
DROP FUNCTION IF EXISTS get_bldn_expense_terms;
DROP FUNCTION IF EXISTS create_service;
DROP FUNCTION IF EXISTS change_service;
DROP FUNCTION IF EXISTS delete_service;
DROP FUNCTION IF EXISTS create_service_mode;
DROP FUNCTION IF EXISTS get_service_modes;
DROP FUNCTION IF EXISTS delete_service_mode;
DROP FUNCTION IF EXISTS change_service_mode;
DROP FUNCTION IF EXISTS get_service_mode;
DROP FUNCTION IF EXISTS get_bldn_services;
DROP FUNCTION IF EXISTS add_bldn_service;
DROP FUNCTION IF EXISTS change_bldn_service;
DROP FUNCTION IF EXISTS delete_bldn_service;
DROP FUNCTION IF EXISTS get_service_in_bldn;
DROP FUNCTION IF EXISTS get_bldn_service_list;
DROP FUNCTION IF EXISTS get_energo_classes;
DROP FUNCTION IF EXISTS get_user;
DROP FUNCTION IF EXISTS adm_get_users;
DROP FUNCTION IF EXISTS adm_create_user;
DROP FUNCTION IF EXISTS adm_change_username;
DROP FUNCTION IF EXISTS adm_change_user_password;
DROP FUNCTION IF EXISTS is_user_valid_password;
DROP FUNCTION IF EXISTS get_user_info;
DROP FUNCTION IF EXISTS get_user_info_by_id;
DROP FUNCTION IF EXISTS get_roles;
DROP FUNCTION IF EXISTS get_user_roles;
DROP FUNCTION IF EXISTS get_user_no_roles;
DROP FUNCTION IF EXISTS adm_has_admin_role;
DROP FUNCTION IF EXISTS adm_add_user_role;
DROP FUNCTION IF EXISTS adm_remove_user_role;
DROP FUNCTION IF EXISTS adm_block_user;
DROP FUNCTION IF EXISTS adm_unblock_user;
DROP FUNCTION IF EXISTS adm_get_access_types;
DROP FUNCTION IF EXISTS adm_create_role;
DROP FUNCTION IF EXISTS adm_role_has_access;
DROP FUNCTION IF EXISTS adm_role_has_no_access;
DROP FUNCTION IF EXISTS adm_add_role_access;
DROP FUNCTION IF EXISTS adm_remove_role_access;
DROP FUNCTION IF EXISTS has_access;
DROP FUNCTION IF EXISTS bldn_address;
DROP FUNCTION IF EXISTS add_log_action;
DROP FUNCTION IF EXISTS log_work_string;
DROP FUNCTION IF EXISTS log_plan_work_string;
DROP FUNCTION IF EXISTS get_bldn_id_no_list;
DROP FUNCTION IF EXISTS get_managed_bldn_id_no_list;
DROP FUNCTION IF EXISTS add_subaccount;
DROP FUNCTION IF EXISTS get_bldn_subaccount;
DROP FUNCTION IF EXISTS get_bldn_subaccounts;
DROP FUNCTION IF EXISTS get_bldn_subaccount_history;
DROP FUNCTION IF EXISTS update_fact_expense;
DROP FUNCTION IF EXISTS get_terms;
DROP FUNCTION IF EXISTS get_services;
DROP FUNCTION IF EXISTS get_service_types;
DROP FUNCTION IF EXISTS plan_expenses_to_gis;
DROP FUNCTION IF EXISTS plan_price_expenses_to_gis;
DROP FUNCTION IF EXISTS get_counter_models;
DROP FUNCTION IF EXISTS get_counter_model;
DROP FUNCTION IF EXISTS create_counter_model;
DROP FUNCTION IF EXISTS change_counter_model;
DROP FUNCTION IF EXISTS delete_counter_model;
DROP FUNCTION IF EXISTS get_bldn_plan_subaccount;
DROP FUNCTION IF EXISTS load_plan_subaccounts;
DROP FUNCTION IF EXISTS get_gwt_round;
DROP FUNCTION IF EXISTS create_maintenance_work;
DROP FUNCTION IF EXISTS get_maintenance_work;
DROP FUNCTION IF EXISTS get_maintenance_work_by_work;
DROP FUNCTION IF EXISTS change_maintenance_work;
DROP FUNCTION IF EXISTS delete_maintenance_work;
DROP FUNCTION IF EXISTS recalc_maintenance_work;
DROP FUNCTION IF EXISTS recalc_maintenance_works;
DROP FUNCTION IF EXISTS get_work_material_type;
DROP FUNCTION IF EXISTS get_work_material_types;
DROP FUNCTION IF EXISTS get_work_materials;
DROP FUNCTION IF EXISTS create_work_material_type;
DROP FUNCTION IF EXISTS change_work_material_type;
DROP FUNCTION IF EXISTS delete_work_material_type;
DROP FUNCTION IF EXISTS add_tmp_counter;
DROP FUNCTION IF EXISTS change_tmp_counter;
DROP FUNCTION IF EXISTS delete_tmp_counter;
DROP FUNCTION IF EXISTS get_all_tmp_counters;
DROP FUNCTION IF EXISTS get_bldn_tmp_counters;
DROP FUNCTION IF EXISTS get_tmp_counter;
DROP FUNCTION IF EXISTS add_counter_certificate;
DROP FUNCTION IF EXISTS get_counter_certificates;
DROP FUNCTION IF EXISTS delete_counter_certificate;
DROP FUNCTION IF EXISTS get_rkc_service;
DROP FUNCTION IF EXISTS get_rkc_services;
DROP FUNCTION IF EXISTS delete_rkc_service;
DROP FUNCTION IF EXISTS create_rkc_service;
DROP FUNCTION IF EXISTS change_rkc_service;
DROP FUNCTION IF EXISTS load_rkc_values;
DROP FUNCTION IF EXISTS get_uk_service;
DROP FUNCTION IF EXISTS get_uk_services;
DROP FUNCTION IF EXISTS delete_uk_service;
DROP FUNCTION IF EXISTS change_uk_service;
DROP FUNCTION IF EXISTS create_uk_service;
DROP FUNCTION IF EXISTS get_bldn_mapping;
DROP FUNCTION IF EXISTS get_bldn_meter_readings;
DROP FUNCTION IF EXISTS load_meter_readings;
DROP FUNCTION IF EXISTS load_subaccounts_sum;
DROP FUNCTION IF EXISTS get_flats_in_term_bldn;
DROP FUNCTION IF EXISTS load_full_flats_info;
DROP FUNCTION IF EXISTS get_flat_terms_in_bldn;
DROP FUNCTION IF EXISTS get_flat_history;
DROP FUNCTION IF EXISTS get_flats_info;
DROP FUNCTION IF EXISTS get_flat_history_info;
DROP FUNCTION IF EXISTS get_added_types;
DROP FUNCTION IF EXISTS load_rkc_addeds;
DROP FUNCTION IF EXISTS bldn_subaccount_percent;
DROP FUNCTION IF EXISTS get_common_property_group;
DROP FUNCTION IF EXISTS get_common_property_group_list;
DROP FUNCTION IF EXISTS create_common_property_group;
DROP FUNCTION IF EXISTS change_common_property_group;
DROP FUNCTION IF EXISTS delete_common_property_group;
DROP FUNCTION IF EXISTS get_common_property_element;
DROP FUNCTION IF EXISTS get_common_property_element_list;
DROP FUNCTION IF EXISTS create_common_property_element;
DROP FUNCTION IF EXISTS change_common_property_element;
DROP FUNCTION IF EXISTS delete_common_property_element;
DROP FUNCTION IF EXISTS get_common_property_element_parameter;
DROP FUNCTION IF EXISTS get_common_property_element_parameters_list;
DROP FUNCTION IF EXISTS create_common_property_element_parameter;
DROP FUNCTION IF EXISTS change_common_property_element_parameter;
DROP FUNCTION IF EXISTS delete_common_property_element_parameter;
DROP FUNCTION IF EXISTS change_bldn_common_property_element_state;
DROP FUNCTION IF EXISTS change_bldn_common_property_element_value;
DROP FUNCTION IF EXISTS change_bldn_common_property_parameter_value;
DROP FUNCTION IF EXISTS get_bldn_common_properties;
DROP FUNCTION IF EXISTS create_and_get_work_annex;
DROP FUNCTION IF EXISTS get_offers_work;
DROP FUNCTION IF EXISTS create_offers_work;
DROP FUNCTION IF EXISTS change_offers_work;
DROP FUNCTION IF EXISTS delete_offers_work;
DROP FUNCTION IF EXISTS get_work_offers_in_bldn;
DROP FUNCTION IF EXISTS get_work_annex_in_bldn;
DROP FUNCTION IF EXISTS get_work_annex;
DROP FUNCTION IF EXISTS load_offers_works;
DROP FUNCTION IF EXISTS load_offers_expenses;
DROP FUNCTION IF EXISTS get_expense_group;
DROP FUNCTION IF EXISTS get_expense_groups;
DROP FUNCTION IF EXISTS create_expense_group;
DROP FUNCTION IF EXISTS change_expense_group;
DROP FUNCTION IF EXISTS delete_expense_group;
DROP FUNCTION IF EXISTS get_man_hour_cost_mode;
DROP FUNCTION IF EXISTS get_man_hour_cost_modes;
DROP FUNCTION IF EXISTS create_man_hour_cost_mode;
DROP FUNCTION IF EXISTS change_man_hour_cost_mode;
DROP FUNCTION IF EXISTS delete_man_hour_cost_mode;
DROP FUNCTION IF EXISTS get_man_hour_cost_sum;
DROP FUNCTION IF EXISTS get_man_hour_cost;
DROP FUNCTION IF EXISTS get_man_hour_cost_rate;
DROP FUNCTION IF EXISTS change_man_hour_cost_sum;
DROP FUNCTION IF EXISTS get_man_hour_cost_by_bldn_term;
DROP FUNCTION IF EXISTS get_man_hour_cost_rates_by_term;
DROP FUNCTION IF EXISTS get_subaccount_terms;
DROP FUNCTION IF EXISTS get_occ_and_address;
DROP FUNCTION IF EXISTS add_man_hour_rates_to_contractor;
DROP FUNCTION IF EXISTS get_accrued_history_by_flat;
DROP FUNCTION IF EXISTS get_bldn_chairman_in_term;
DROP FUNCTION IF EXISTS add_signature;
DROP FUNCTION IF EXISTS get_signature_in_term;
DROP FUNCTION IF EXISTS delete_chairman_signature;
DROP FUNCTION IF EXISTS get_chairmans_signature_in_bldn;
DROP FUNCTION IF EXISTS get_employee_signature;

DROP PROCEDURE IF EXISTS write_action_to_log;

DROP FUNCTION IF EXISTS get_file_version;
DROP FUNCTION IF EXISTS is_not_value;
DROP FUNCTION IF EXISTS is_all_values;
DROP FUNCTION IF EXISTS get_not_value;
DROP FUNCTION IF EXISTS get_error_number;
DROP FUNCTION IF EXISTS get_error_message;
DROP FUNCTION IF EXISTS rights_get_counter_rights_number;
DROP FUNCTION IF EXISTS rights_get_certificate_rights_number;
DROP FUNCTION IF EXISTS rights_get_works_rights_number;
DROP FUNCTION IF EXISTS rights_get_plan_rights_number;
DROP FUNCTION IF EXISTS rights_get_bldn_accrued_rights_number;
DROP FUNCTION IF EXISTS rights_get_contractor_cost_rights_number;
DROP FUNCTION IF EXISTS rights_get_meter_readings_rights_number;
DROP FUNCTION IF EXISTS rights_get_owners_rights_number;
DROP FUNCTION IF EXISTS rights_get_common_property_elements_rights_number;
DROP FUNCTION IF EXISTS rights_get_common_property_group_rights_number;
DROP FUNCTION IF EXISTS rights_get_offers_work_rights_number;
DROP FUNCTION IF EXISTS rights_get_tech_info_rights_number;

DROP FUNCTION IF EXISTS user_has_right_read;
DROP FUNCTION IF EXISTS user_has_right_change;
DROP FUNCTION IF EXISTS user_has_right_delete;
DROP TRIGGER IF EXISTS access_rights__insert__trigger ON access_rights;
DROP FUNCTION IF EXISTS add_access_right_trigger;
DROP TRIGGER IF EXISTS access_rights__delete__trigger ON access_rights;
DROP FUNCTION IF EXISTS delete_access_right_trigger;
DROP TRIGGER IF EXISTS roles__insert__trigger ON roles;
DROP FUNCTION IF EXISTS add_role_trigger;
DROP TRIGGER IF EXISTS roles__delete__trigger ON roles;
DROP FUNCTION IF EXISTS delete_role_trigger;
DROP TRIGGER IF EXISTS gwt__insert__trigger ON global_work_types;
DROP FUNCTION IF EXISTS add_gwt_trigger;
DROP TRIGGER IF EXISTS gwt__delete__trigger ON global_work_types;
DROP FUNCTION IF EXISTS delete_gwt_trigger;


DROP TRIGGER IF EXISTS buildings__insert__index ON buildings;
DROP FUNCTION IF EXISTS add_building;
DROP TRIGGER IF EXISTS buildings__delete__trigger ON buildings;
DROP FUNCTION IF EXISTS drop_building;

DROP FUNCTION IF EXISTS report_1;
DROP FUNCTION IF EXISTS report_2;
DROP FUNCTION IF EXISTS report_3;
DROP FUNCTION IF EXISTS report_4;
DROP FUNCTION IF EXISTS report_5;
DROP FUNCTION IF EXISTS report_6;
DROP FUNCTION IF EXISTS report_7;
DROP FUNCTION IF EXISTS report_8;
DROP FUNCTION IF EXISTS report_9;
DROP FUNCTION IF EXISTS report_10;
DROP FUNCTION IF EXISTS bldnPassport;
DROP FUNCTION IF EXISTS all_works;
DROP FUNCTION IF EXISTS sub_accounts;
DROP FUNCTION IF EXISTS report_bldn_common_properties;
DROP FUNCTION IF EXISTS report_year_plan;
DROP FUNCTION IF EXISTS report_mainworkmaterials;
DROP FUNCTION IF EXISTS report_contractormaterials;
DROP FUNCTION IF EXISTS report_101;
DROP FUNCTION IF EXISTS report_101a;
DROP FUNCTION IF EXISTS report_102;
DROP FUNCTION IF EXISTS report_11;
DROP FUNCTION IF EXISTS report_12;
DROP FUNCTION IF EXISTS report_13;
DROP FUNCTION IF EXISTS report_14;
DROP FUNCTION IF EXISTS report_201;
DROP FUNCTION IF EXISTS report_bldn_work_completition;

-- тип для информации по квартирам
DROP TYPE IF EXISTS flats_info;
CREATE TYPE flats_info AS (
  bldn_id INTEGER
  , term_id INTEGER
  , flat_id BIGINT
  , flat_no VARCHAR
  , residental BOOLEAN
  , uninhabitable BOOLEAN
  , rooms INTEGER
  , passport_square NUMERIC
  , square NUMERIC
  , note TEXT
  , cadastral_no TEXT
  , share_numerator INTEGER
  , share_denominator INTEGER
  , is_legal_entity BOOLEAN
  , is_privatized BOOLEAN
  , id BIGINT
  , owner_name VARCHAR
  , owner_document VARCHAR
  , phone VARCHAR
  , has_pd_consent BOOLEAN
  , is_chairman BOOLEAN
  , is_sekretar BOOLEAN
  , is_senat BOOLEAN
  , saldo NUMERIC
  , sort_flat_no TEXT
);

DROP VIEW IF EXISTS streetlist;
DROP VIEW IF EXISTS contractorlist;
DROP VIEW IF EXISTS xstreets;
DROP VIEW IF EXISTS buildings_workslist;
DROP VIEW IF EXISTS bldn_expenses;
DROP VIEW IF EXISTS managed_buildings;
DROP VIEW IF EXISTS bldn_id_no_list;
DROP VIEW IF EXISTS current_subaccounts;
DROP VIEW IF EXISTS maintenance_works;
DROP VIEW IF EXISTS certificate_tmp_counters;
DROP VIEW IF EXISTS common_property_dictionary;

DROP TABLE IF EXISTS chairman_signature;
DROP TABLE IF EXISTS lawsuit_persons;
DROP TABLE IF EXISTS lawsuits;
DROP TABLE IF EXISTS bldn_man_hour_cost;
DROP TABLE IF EXISTS man_hour_cost_rates;
DROP TABLE IF EXISTS building_common_property_elements_history;
DROP TABLE IF EXISTS building_common_property_element_parameter_history;
DROP TABLE IF EXISTS building_common_property_element_parameter;
DROP TABLE IF EXISTS building_common_property_elements;
DROP TABLE IF EXISTS common_property_element_parameter;
DROP TABLE IF EXISTS plan_subaccounts;
DROP TABLE IF EXISTS bldn_subaccounts;
DROP TABLE IF EXISTS expense_groups;
DROP TABLE IF EXISTS bldn_expense_names;
DROP TABLE IF EXISTS expenses;
DROP TABLE IF EXISTS expense_items;
DROP TABLE IF EXISTS bldn_services_history;
DROP TABLE IF EXISTS bldn_services;
DROP TABLE IF EXISTS service_modes;
DROP TABLE IF EXISTS services;
DROP TABLE IF EXISTS old_works;
DROP TABLE IF EXISTS plan_works;
DROP TABLE IF EXISTS plan_work_statuses;
DROP TABLE IF EXISTS works;
DROP TABLE IF EXISTS work_financing_sources;
DROP TABLE IF EXISTS accrueds;
DROP TABLE IF EXISTS buildings_history;
DROP TABLE IF EXISTS buildings_tech_info_history;
DROP TABLE IF EXISTS buildings_land_info_history;
DROP TABLE IF EXISTS buildings_land_info;
DROP TABLE IF EXISTS buildings_tech_info;
DROP TABLE IF EXISTS buildings;
DROP TABLE IF EXISTS bldn_types;
DROP TABLE IF EXISTS dogovors;
DROP TABLE IF EXISTS constants;
DROP TABLE IF EXISTS wall_materials;
DROP TABLE IF EXISTS terms;
DROP TABLE IF EXISTS work_kinds;
DROP TABLE IF EXISTS work_types;
DROP TABLE IF EXISTS global_work_types;
DROP TABLE IF EXISTS streets;
DROP TABLE IF EXISTS street_types;
DROP TABLE IF EXISTS villages;
DROP TABLE IF EXISTS village_types;
DROP TABLE IF EXISTS municipal_districts;
DROP TABLE IF EXISTS improvements;
DROP TABLE IF EXISTS contractors;
DROP TABLE IF EXISTS employees;
DROP TABLE IF EXISTS position_statuses;
DROP TABLE IF EXISTS management_companies;
DROP TABLE IF EXISTS log_log;
DROP TABLE IF EXISTS log_operations;
DROP TABLE IF EXISTS roles_access;
DROP TABLE IF EXISTS user_roles;
DROP TABLE IF EXISTS access_types;
DROP TABLE IF EXISTS roles;
DROP TABLE IF EXISTS users;
DROP TABLE IF EXISTS energo_classes;
DROP TABLE IF EXISTS service_types;
DROP TABLE IF EXISTS counter_models;
DROP TABLE IF EXISTS boolean_access_names;
DROP TABLE IF EXISTS hidden_maintenance_works;
DROP TABLE IF EXISTS works_materials;
DROP TABLE IF EXISTS work_material_types;
DROP TABLE IF EXISTS tmp_counters;
DROP TABLE IF EXISTS certificates;
DROP TABLE IF EXISTS rkc_values_history;
DROP TABLE IF EXISTS rkc_services;
DROP TABLE IF EXISTS uk_services;
DROP TABLE IF EXISTS uk_accrued_source;
DROP TABLE IF EXISTS bldn_id_mapping;
DROP TABLE IF EXISTS meter_readings;
DROP TABLE IF EXISTS meter_readings_tmp;
DROP TABLE IF EXISTS owners;
DROP TABLE IF EXISTS flat_shares;
DROP TABLE IF EXISTS flats
DROP TABLE IF EXISTS rkc_addeds_history;
DROP TABLE IF EXISTS rkc_added_types;
DROP TABLE IF EXISTS common_property_element;
DROP TABLE IF EXISTS common_property_group;
DROP TABLE IF EXISTS offers_annex;
DROP TABLE IF EXISTS offers_work;
DROP TABLE IF EXISTS offers_expense;
DROP TABLE IF EXISTS log_offers;
DROP TABLE IF EXISTS man_hour_cost_modes;

DROP TABLE IF EXISTS roles_access_rights;
DROP TABLE IF EXISTS access_rights;
DROP TABLE IF EXISTS errors;
DROP TABLE IF EXISTS load_files_versions;

CREATE TABLE boolean_access_names (id INTEGER, name VARCHAR(200));
INSERT INTO boolean_access_names VALUES (1, 'Есть доступ');

CREATE TABLE access_rights (
       id SERIAL PRIMARY KEY,
       name TEXT NOT NULL UNIQUE
);
COMMENT ON TABLE access_rights IS 'Права доступа';
COMMENT ON COLUMN access_rights.id IS 'код права доступа';
COMMENT ON COLUMN access_rights.name IS 'название права доступа';

INSERT INTO access_rights(id, name) VALUES
(1, 'Акты ОДПУ')
, (2, 'Приборы учёта')
, (3, 'Начисления услуг')
, (4, 'Расходы подрядчиков')
, (5, 'Показания ИПУ')
, (6, 'Собственники помещений')
, (7, 'Элементы общего имущества')
, (8, 'Предлагаемые работы')
, (9, 'Техническая информация')
;

CREATE TABLE users
       (id SERIAL PRIMARY KEY,
       login VARCHAR(20),
       name VARCHAR(200),
       password TEXT,
       is_active BOOLEAN DEFAULT TRUE);
INSERT INTO users(id, login, name, password) VALUES (0, 'admin', 'admin', crypt('admin', gen_salt('md5')));

CREATE TABLE roles
       (id SERIAL PRIMARY KEY,
       name VARCHAR(100));
INSERT INTO roles(id, name) VALUES (1, 'Администратор');
SELECT SETVAL('roles_id_seq', MAX(id)) FROM roles;

CREATE TABLE roles_access_rights (
       access_id INTEGER REFERENCES access_rights ON DELETE CASCADE ON UPDATE CASCADE,
       role_id INTEGER REFERENCES roles ON DELETE CASCADE ON UPDATE CASCADE,
       access_read BOOLEAN NOT NULL DEFAULT FALSE,
       access_change BOOLEAN NOT NULL DEFAULT FALSE,
       access_delete BOOLEAN NOT NULL DEFAULT FALSE,
       CONSTRAINT role_access_must_be_unique UNIQUE (access_id, role_id)
);

CREATE TABLE user_roles
       (user_id INTEGER NOT NULL REFERENCES users ON DELETE CASCADE ON UPDATE CASCADE,
       role_id INTEGER NOT NULL REFERENCES roles ON DELETE CASCADE ON UPDATE CASCADE);
INSERT INTO user_roles VALUES (0, 1);

CREATE TABLE access_types
       (id SERIAL PRIMARY KEY,
       acs_type ACCESS_TYPE NOT NULL UNIQUE,
       name VARCHAR(100) NOT NULL);
INSERT INTO access_types(acs_type, name) VALUES ('gwt', 'Виды ремnонтов');

CREATE TABLE roles_access
       (role_id INTEGER NOT NULL REFERENCES roles ON DELETE CASCADE ON UPDATE CASCADE,
       acs_id INTEGER NOT NULL REFERENCES access_types ON DELETE CASCADE ON UPDATE CASCADE,
       acs_value INTEGER);

CREATE TABLE log_operations (
       id SERIAL PRIMARY KEY,
       name VARCHAR(200) NOT NULL);
COMMENT ON TABLE log_operations IS 'Действия с базой';
COMMENT ON COLUMN log_operations.id IS 'код действия';
COMMENT ON COLUMN log_operations.name IS 'название действия';
INSERT INTO log_operations (id, name) VALUES
(1, 'Добавление работы'),
(2, 'Изменение работы'),
(3, 'Удаление работы'),
(4, 'Добавление планируемой работы'),
(5, 'Изменение планируемой работы'),
(6, 'Удаление планируемой работы'),
(7, 'Создание группы структуры платы'),
(8, 'Изменение группы структуры платы'),
(9, 'Удаление группы структуры платы'),
(10, 'Создание статьи расходов'),
(11, 'Изменение статьи расходов'),
(12, 'Удаление статьи расходов'),
(13, 'Создание группы элементов общего имущества'),
(14, 'Изменение группы элементов общего имущества'),
(15, 'Удаление группы элементов общего имущества'),
(16, 'Создание параметра элемента общего имущества дома'),
(17, 'Изменение параметра элемента общего имущества дома'),
(18, 'Удаление парамента элемента общего имущества дома'),
(19, 'Создание элемента общего имущества дома'),
(20, 'Изменение элемента общего имущества дома'),
(21, 'Удаление элемента общего имущества дома'),
(22, 'Изменение элемента общего имущества в доме'),
(23, 'Изменение параметра элемента общего имущества в доме'),
(24, 'Создание режима стоимости человекочаса'),
(25, 'Изменение режима стоимости человекочаса'),
(26, 'Удаление режима стоимости человекочаса'),
(27, 'Добавление работы по содержанию'),
(28, 'Изменение работы по содержанию'),
(29, 'Удаление работы по содержанию'),
(30, 'Изменение режима человекочаса в доме'),
(31, 'Добавление подрядной организации'),
(32, 'Изменение подрядной организации'),
(33, 'Удаление подрядной организации'),
(34, 'Изменение стоимости человекочаса'),
(35, 'Добавление ЖКУ'),
(36, 'Изменение ЖКУ'),
(37, 'Удаление ЖКУ'),
(38, 'Добавление режима ЖКУ'),
(39, 'Изменение режима ЖКУ'),
(40, 'Удаление режима ЖКУ'),
(100, 'Создание дома'),
(101, 'Удаление дома'),
(102, 'Изменение ЖКУ дома'),
(103, 'Изменение общей информации дома'),
(104, 'Изменение информации о договорах дома'),
(105, 'Изменение технической информации МКД'),
(106, 'Загрузка помещений, долей, собственников'),
(107, 'Добавление подписи коменданта'),
(108, 'Удаление подписи коменданта');

CREATE TABLE log_log (
  action_id INTEGER NOT NULL REFERENCES log_operations ON DELETE RESTRICT ON UPDATE CASCADE
  , action_time TIMESTAMP DEFAULT LOCALTIMESTAMP
  , user_id INTEGER NOT NULL REFERENCES users ON DELETE RESTRICT ON UPDATE CASCADE
  , pc_name VARCHAR(100) NOT NULL
  , action_description TEXT
  , log_action JSONB
);
COMMENT ON TABLE log_log IS 'История изменений работ';
COMMENT ON COLUMN log_log.action_id IS 'Код вида операции';
COMMENT ON COLUMN log_log.action_time IS 'Время операции';
COMMENT ON COLUMN log_log.action_description IS 'Описание операции';
COMMENT ON COLUMN log_log.user_id IS 'Код пользователя';
COMMENT ON COLUMN log_log.pc_name IS 'Название компьютера';
COMMENT ON COLUMN log_log.log_action IS 'Описание операции';

CREATE TABLE load_files_versions (
  id VARCHAR(20) PRIMARY KEY
  , name VARCHAR(100) NOT NULL
  , file_version INTEGER NOT NULL
);
COMMENT ON TABLE load_files_versions IS 'Версии загружаемых xml-файлов';
COMMENT ON COLUMN load_files_versions.id IS 'код (первичный ключ)';
COMMENT ON COLUMN load_files_versions.name IS 'наименование';
COMMENT ON COLUMN load_files_versions.file_version IS 'версия файла';
INSERT INTO load_files_versions VALUES
				  ('full_flat', 'Информация о помещениях', 1),
				  ('avr', 'Информация по подрядчикам', 1),
				  ('flat_square', 'Информация о площадях', 1),
				  ('expense', 'Информация о структуре', 1),
				  ('subaccount', 'Информация о субсчетах', 1),
				  ('plan_subaccount', 'Информация о плановых субсчетах', 1),
				  ('month_subaccount', 'Информация о субсчетах за месяц', 1),
				  ('rkc_addeds', 'Разовые начисления РКЦ', 1),
				  ('offers', 'Предложения к договорам', 1);

CREATE TABLE constants
       (name VARCHAR(50) UNIQUE,
       value VARCHAR(20),
       note TEXT);
INSERT INTO constants VALUES
			('version', '2.0.4', 'Версия программы'),
			('gwt_work_round', 1, 'Количество знаков после запятой для работ по содержанию'),
			('all_values', -1002, 'Значение, передаваемое при выборе пункта "Все"'),
			('not_value', -1001, 'Значение, передаваемое как флаг непереданного значения'),
			('gwt_work_access_prefix', 1000000, 'Множитель для проверки доступа к работам'),
			('gwt_planwork_access_prefix', 1100000, 'Множитель для проверки доступа к планам работ')
			, ('common_property_group_prefix', 2000000, 'Множитель для проверки доступа к группам элементов общего имущества дома');

CREATE TABLE errors(
  name VARCHAR(30),
  err_number VARCHAR(20),
  message TEXT);
COMMENT ON TABLE errors IS 'Ошибки программы';
COMMENT ON COLUMN errors.name IS 'Название ошибки';
COMMENT ON COLUMN errors.err_number IS 'Номер ошибки';
COMMENT ON COLUMN errors.message IS 'Сообщение об ошибке';
INSERT INTO errors VALUES
		     ('has_no_access', '60010', 'Не хватает прав'),
		     ('has_no_values', '60002', 'Нет данных'),
		     ('file_version_error', '60011', 'Несоответствие версии файла загрузки'),
		     ('delete_denied', '99003', 'Нельзя удалять'),
		     ('has_children', '60003', 'Имеются ссылки');


CREATE TABLE energo_classes (
       id SERIAL PRIMARY KEY,
       name VARCHAR(20)
);
INSERT INTO energo_classes VALUES (0, 'Не указано');
INSERT INTO energo_classes(name) VALUES ('A++'), ('A+'), ('A'), ('B'), ('C'), ('D'), ('E'), ('F'), ('G');

CREATE TABLE management_companies
       (id SERIAL PRIMARY KEY,
       name VARCHAR(100) NOT NULL CONSTRAINT mc_name_must_be_unique UNIQUE,
       report_name VARCHAR(20),
       not_manage BOOLEAN);

INSERT INTO management_companies VALUES (0, 'Непосредственное управление', 'Частный сектор', TRUE);

CREATE TABLE position_statuses
       (id SERIAL PRIMARY KEY,
       name VARCHAR(100) NOT NULL);

INSERT INTO position_statuses VALUES (1, 'Директор'), (2, 'Главный инженер'), (3, 'Другое');

SELECT SETVAL('position_statuses_id_seq', max(id)) FROM position_statuses;

CREATE TABLE employees
       (id SERIAL PRIMARY KEY,
       organization_id INTEGER NOT NULL REFERENCES management_companies ON DELETE CASCADE ON UPDATE CASCADE,
       last_name VARCHAR(60) NOT NULL,
       first_name VARCHAR(60),
       second_name VARCHAR(60),
       position_status INTEGER REFERENCES position_statuses ON DELETE RESTRICT ON UPDATE CASCADE,
       position_name VARCHAR(100) NOT NULL,
       sign_report BOOLEAN,
       signature BYTEA);
COMMENT ON COLUMN employees.signature IS 'Подпись';

CREATE TABLE contractors
       (id SERIAL PRIMARY KEY,
       name VARCHAR(100) NOT NULL CONSTRAINT contractor_name_must_be_unique UNIQUE,
       director VARCHAR(200),
       director_position VARCHAR(200),
       bldn_contractor BOOLEAN,
       is_using BOOLEAN DEFAULT TRUE);
COMMENT ON TABLE contractors IS 'Подрядные организации';
COMMENT ON COLUMN contractors.id  IS 'Код';
COMMENT ON COLUMN contractors.name IS 'Название';
COMMENT ON COLUMN contractors.director IS 'Директор';
COMMENT ON COLUMN contractors.director_position IS 'Должность директора';
COMMENT ON COLUMN contractors.bldn_contractor IS 'Обслуживает дома';
COMMENT ON COLUMN contractors.is_using IS 'Активен';
INSERT INTO contractors VALUES (0, 'Не указан', NULL, NULL, True, FALSE);

CREATE TABLE improvements
       (id SERIAL PRIMARY KEY,
       name VARCHAR(100) NOT NULL,
       short_name VARCHAR(30));

INSERT INTO improvements VALUES (0, 'Не указана', 'Не указана');

CREATE TABLE municipal_districts
       (id SERIAL PRIMARY KEY,
       name VARCHAR(100) NOT NULL CONSTRAINT md_name_must_be_unique UNIQUE,
       head VARCHAR(200) NOT NULL,
       head_position VARCHAR(300));

CREATE TABLE village_types
       (id SERIAL PRIMARY KEY,
       name VARCHAR(20) NOT NULL,
       short_name VARCHAR(10));

CREATE TABLE villages
       (id SERIAL PRIMARY KEY,
       md_id INTEGER NOT NULL REFERENCES municipal_districts ON DELETE RESTRICT ON UPDATE CASCADE,
       name VARCHAR(100) NOT NULL,
       site_name VARCHAR(100),
       CONSTRAINT village_in_md_must_be_unique UNIQUE (name, md_id));

CREATE TABLE street_types
       (id SERIAL PRIMARY KEY,
       name VARCHAR(20) NOT NULL,
       short_name VARCHAR(10));

INSERT INTO street_types VALUES (0, 'нет', ''), (1, 'улица', 'ул.'), (2, 'переулок', 'пер.'), (3, 'площадь', 'пл.');

SELECT SETVAL('street_types_id_seq', max(id)) FROM street_types;

CREATE TABLE streets
       (id SERIAL PRIMARY KEY,
       village_id INTEGER NOT NULL REFERENCES villages ON DELETE RESTRICT ON UPDATE CASCADE,
       name VARCHAR(100),
       site_name VARCHAR(100),
       street_type INTEGER DEFAULT(0) REFERENCES street_types ON DELETE RESTRICT ON UPDATE CASCADE,
       CONSTRAINT street_in_village_must_be_unique UNIQUE (name, village_id, street_type));

CREATE TABLE global_work_types
       (id SERIAL PRIMARY KEY,
       name VARCHAR(50) NOT NULL CONSTRAINT gwt_name_must_be_unique UNIQUE,
       description TEXT);

INSERT INTO global_work_types VALUES (1, 'Содержание',''), (2, 'Текущий ремонт', ''), (3, 'Капитальный ремонт', '');

SELECT SETVAL('global_work_types_id_seq', max(id)) FROM global_work_types;

CREATE TABLE work_types
       (id SERIAL PRIMARY KEY,
       name VARCHAR(100) NOT NULL);

CREATE TABLE work_kinds
       (id SERIAL PRIMARY KEY,
       worktype_id INTEGER NOT NULL REFERENCES work_types ON DELETE RESTRICT ON UPDATE CASCADE,
       name VARCHAR(300) NOT NULL,
       CONSTRAINT work_kind_in_work_type_must_be_unique UNIQUE (name, worktype_id));
       	     		   
CREATE TABLE terms
       (id SERIAL PRIMARY KEY,
       begin_date DATE,
       end_date DATE);

CREATE TABLE wall_materials
       (id SERIAL PRIMARY KEY,
       name VARCHAR(100));

INSERT INTO wall_materials VALUES (0, 'Не указано');

       
CREATE TABLE dogovors
       (id SERIAL PRIMARY KEY,
       name VARCHAR(100),
       short_name VARCHAR(30));

INSERT INTO dogovors VALUES (0, 'Не заключен', 'не заключен');

CREATE TABLE service_types (
       id SERIAL PRIMARY KEY,
       type_name VARCHAR(200) NOT NULL);
COMMENT ON TABLE service_types IS 'Типы ЖКУ';
COMMENT ON COLUMN service_types.id IS 'Код типа ЖКУ';
COMMENT ON COLUMN service_types.type_name IS 'Название типа ЖКУ';

INSERT INTO service_types (id, type_name) VALUES (0, 'Прочее'), (1, 'Отопление'), (2, 'Горячее водоснабжение'), (3, 'Газоснабжение'), (4, 'Электроснабжение');

CREATE TABLE services
       (id INTEGER PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY,
       service_type INTEGER NOT NULL DEFAULT 0 REFERENCES service_types ON DELETE SET DEFAULT ON UPDATE CASCADE,
       name VARCHAR(100) NOT NULL CONSTRAINT service_name_must_be_unique UNIQUE,
       is_print_to_passport BOOLEAN DEFAULT FALSE
);
COMMENT ON TABLE services IS 'Услуги в домах';
COMMENT ON COLUMN services.id IS 'Код услуги';
COMMENT ON COLUMN services.service_type IS 'Код типа услуги';
COMMENT ON COLUMN services.name IS 'Название услуги';
COMMENT ON COLUMN  services.is_print_to_passport IS 'Признак вывода услуги в паспорт подготовки к зиме';

CREATE TABLE service_modes
       (id SERIAL PRIMARY KEY,
       service_id INTEGER REFERENCES services ON DELETE CASCADE ON UPDATE CASCADE,
       mode_name VARCHAR(200) NOT NULL);
COMMENT ON TABLE service_modes IS 'Режимы ЖКУ';
COMMENT ON COLUMN service_modes.id IS 'Код режима';
COMMENT ON COLUMN service_modes.service_id IS 'Код услуги';
COMMENT ON COLUMN service_modes.mode_name IS 'Название режима';

CREATE TABLE bldn_types (
     id SERIAL PRIMARY KEY,
     name VARCHAR(50) NOT NULL
);

INSERT INTO bldn_types VALUES (1, 'Многоквартирый'), (2, 'Жилой'), (3, 'Блокированной застройки');

CREATE TABLE buildings
       (id SERIAL PRIMARY KEY,
       street_id INTEGER REFERENCES streets ON DELETE RESTRICT ON UPDATE CASCADE,
       bldn_no VARCHAR(10),
       mc_id INTEGER DEFAULT 0 REFERENCES management_companies ON DELETE RESTRICT ON UPDATE CASCADE,
       contractor_id INTEGER DEFAULT 0 REFERENCES contractors ON DELETE RESTRICT ON UPDATE CASCADE,
       improvement_id INTEGER DEFAULT 0 REFERENCES improvements ON DELETE RESTRICT ON UPDATE CASCADE,
       cadastral_no VARCHAR(20),
       out_report BOOLEAN DEFAULT FALSE,
       hot_water INTEGER DEFAULT 0,
       cold_water INTEGER DEFAULT 0,
       heating INTEGER DEFAULT 0,
       gas INTEGER DEFAULT 0,
       dogovor_type INTEGER DEFAULT 0 REFERENCES dogovors ON DELETE RESTRICT ON UPDATE CASCADE,
       site_no VARCHAR(10),
       bldn_type INTEGER NOT NULL DEFAULT 1 REFERENCES bldn_types ON DELETE RESTRICT ON UPDATE CASCADE,
       disrepair BOOLEAN NOT NULL DEFAULT FALSE,
       energo_class INTEGER NOT NULL DEFAULT 0 REFERENCES energo_classes ON DELETE SET DEFAULT ON UPDATE CASCADE,
       contract_no VARCHAR(10),
       contract_date DATE,
       fias VARCHAR(36),
       gis_guid VARCHAR(50),
       CONSTRAINT bldn_no_must_be_unique UNIQUE(bldn_no, street_id));

COMMENT ON TABLE buildings IS 'Общая информация о зданиях';
COMMENT ON COLUMN buildings.id IS 'Код дома';
COMMENT ON COLUMN buildings.bldn_no IS 'Номер дома';
COMMENT ON COLUMN buildings.street_id IS 'Код улицы';
COMMENT ON COLUMN buildings.mc_id IS 'Код управляющей компании дома';
COMMENT ON COLUMN buildings.contractor_id IS 'Код подрядной организации, обслуживающей дом';
COMMENT ON COLUMN buildings.improvement_id IS 'Код степени благоустройства';
COMMENT ON COLUMN buildings.cadastral_no IS 'Кадастровый номер здания';
COMMENT ON COLUMN buildings.out_report IS 'Признак вывода отчёта по дому';
COMMENT ON COLUMN buildings.hot_water IS 'Вид горячегоо водоснабжения';
COMMENT ON COLUMN buildings.cold_water IS 'Вид холодного водоснабжения';
COMMENT ON COLUMN buildings.heating IS 'Вид отопления';
COMMENT ON COLUMN buildings.gas IS 'Вид газоснабжения';
COMMENT ON COLUMN buildings.dogovor_type IS 'Код вида договора с УК';
COMMENT ON COLUMN buildings.site_no IS 'Номер дома для формирования отчетов для сайта';
COMMENT ON COLUMN buildings.contract_no IS 'Номер договора с УК';
COMMENT ON COLUMN buildings.contract_date IS 'Дата договора с УК';
COMMENT ON COLUMN buildings.bldn_type IS 'Код типа здания (МКД, частный, ..)';
COMMENT ON COLUMN buildings.disrepair IS 'Признак аварийности';
COMMENT ON COLUMN buildings.energo_class IS 'Код класса энергосбережения';
COMMENT ON COLUMN buildings.fias IS 'Код ФИАС';
COMMENT ON COLUMN buildings.gis_guid IS 'Код дома в ГИС ЖКХ';

CREATE TABLE buildings_tech_info
       (bldn_id INTEGER PRIMARY KEY REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE,
       floor_min INTEGER CONSTRAINT positive_floor_min CHECK (floor_min >= 0), 
       floor_max INTEGER CONSTRAINT floor_max_more_floor_min CHECK (floor_max >= floor_min),
       vaults INTEGER,
       entrances INTEGER,
       stairs INTEGER DEFAULT(NULL),
       built_year INTEGER,
       commissioning_year INTEGER,
       depreciation REAL,
       attic_square REAL,
       vaults_square REAL,
       stairs_square REAL,
       corridor_square REAL,
       other_square REAL,
       structural_volume REAL,
       wallmater_id INTEGER DEFAULT 0 REFERENCES wall_materials ON DELETE RESTRICT ON UPDATE CASCADE,
       has_odpu_electro BOOLEAN DEFAULT FALSE,
       has_odpu_hotwater BOOLEAN DEFAULT FALSE,
       has_odpu_common BOOLEAN DEFAULT FALSE,
       has_odpu_heating BOOLEAN DEFAULT FALSE,
       has_odpu_coldwater BOOLEAN DEFAULT FALSE,
       has_doorphone BOOLEAN DEFAULT FALSE,
       doorphone_comment TEXT,
       has_thermoregulator BOOLEAN DEFAULT FALSE,
       has_doorcloser BOOLEAN DEFAULT FALSE,
       square_banisters REAL,
       square_doors REAL,
       square_windowsills REAL,
       square_doorhandles REAL,
       square_mailboxes REAL,
       square_radiators REAL
);

COMMENT ON TABLE buildings_tech_info IS 'Технические характеристики зданий';
COMMENT ON COLUMN buildings_tech_info.bldn_id IS 'Код дома';
COMMENT ON COLUMN buildings_tech_info.floor_min IS 'Количество этажей минимальное';
COMMENT ON COLUMN buildings_tech_info.floor_max IS 'Количество этажей максимальное';
COMMENT ON COLUMN buildings_tech_info.vaults IS 'Количество подвалов';
COMMENT ON COLUMN buildings_tech_info.entrances IS 'Количество подъездов';
COMMENT ON COLUMN buildings_tech_info.stairs IS 'Количество лестниц';
COMMENT ON COLUMN buildings_tech_info.built_year IS 'Год постройки';
COMMENT ON COLUMN buildings_tech_info.commissioning_year IS 'Год ввода в эксплуатацию';
COMMENT ON COLUMN buildings_tech_info.depreciation IS 'Износ';
COMMENT ON COLUMN buildings_tech_info.attic_square IS 'Площадь чердаков';
COMMENT ON COLUMN buildings_tech_info.vaults_square IS 'Площадь подвалов';
COMMENT ON COLUMN buildings_tech_info.stairs_square IS 'Площадь лестничных площадок и маршев';
COMMENT ON COLUMN buildings_tech_info.corridor_square IS 'Площадь коридоров МОП';
COMMENT ON COLUMN buildings_tech_info.other_square IS 'Площадь иных помещений МОП';
COMMENT ON COLUMN buildings_tech_info.structural_volume IS 'Строительный объем';
COMMENT ON COLUMN buildings_tech_info.wallmater_id IS 'Код материала стен';
COMMENT ON COLUMN buildings_tech_info.has_odpu_electro IS 'Наличие ОДПУ электроэнергии';
COMMENT ON COLUMN buildings_tech_info.has_odpu_hotwater IS 'Наличие ОДПУ горячего водоснабжения';
COMMENT ON COLUMN buildings_tech_info.has_odpu_common IS 'Наличие ОДПУ общей системы теплоснабжения';
COMMENT ON COLUMN buildings_tech_info.has_odpu_heating IS 'Наличие ОДПУ отопления';
COMMENT ON COLUMN buildings_tech_info.has_odpu_coldwater IS 'Наличие ОДПУ холодного водоснабжения';
COMMENT ON COLUMN buildings_tech_info.has_doorphone IS 'Наличие домофонов';
COMMENT ON COLUMN buildings_tech_info.doorphone_comment IS 'Примечание к домофонам';
COMMENT ON COLUMN buildings_tech_info.has_thermoregulator IS 'Наличие погодозависимой автоматики';
COMMENT ON COLUMN buildings_tech_info.square_banisters IS 'Площадь перил';
COMMENT ON COLUMN buildings_tech_info.square_doors IS 'Площадь дверей';
COMMENT ON COLUMN buildings_tech_info.square_windowsills IS 'Площадь подоконников';
COMMENT ON COLUMN buildings_tech_info.square_doorhandles IS 'Площадь ручек';
COMMENT ON COLUMN buildings_tech_info.square_mailboxes IS 'Площадь почтовых ящиков';
COMMENT ON COLUMN buildings_tech_info.square_radiators IS 'Площадь радиаторов МОП';
COMMENT ON COLUMN buildings_tech_info.has_doorcloser IS 'Наличие доводчика';

CREATE TABLE buildings_land_info
       (bldn_id INTEGER PRIMARY KEY REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE,
       inventory_area REAL,
       use_area REAL,
       survey_area REAL,
       builtup_area REAL,
       undeveloped_area REAL,
       hard_coatings REAL,
       drive_ways_hard REAL,
       side_walks_hard REAL,
       others_hard REAL,
       cadastral_no VARCHAR(20),
       saf BOOLEAN DEFAULT FALSE,
       fences BOOLEAN DEFAULT FALSE,
       benches INTEGER);

COMMENT ON TABLE buildings_land_info IS 'Характеристики земельных участков зданий';
COMMENT ON COLUMN buildings_land_info.bldn_id IS 'Код дома';
COMMENT ON COLUMN buildings_land_info.inventory_area IS 'Площадь по данным технической инвентаризации';
COMMENT ON COLUMN buildings_land_info.use_area IS 'Площадь по фактическому использованию';
COMMENT ON COLUMN buildings_land_info.survey_area IS 'Площадь по данным межевания';
COMMENT ON COLUMN buildings_land_info.builtup_area IS 'Площадь застройки';
COMMENT ON COLUMN buildings_land_info.undeveloped_area IS 'Незастроенная площадь';
COMMENT ON COLUMN buildings_land_info.hard_coatings IS 'Всего твердые покрытия';
COMMENT ON COLUMN buildings_land_info.drive_ways_hard IS 'Площадь проездов';
COMMENT ON COLUMN buildings_land_info.side_walks_hard IS 'Площадь тротуаров';
COMMENT ON COLUMN buildings_land_info.others_hard IS 'Прочие твердые покрытия';
COMMENT ON COLUMN buildings_land_info.cadastral_no IS 'Кадастровый номер';
COMMENT ON COLUMN buildings_land_info.saf IS 'Наличие малых архитектурных форм';
COMMENT ON COLUMN buildings_land_info.fences IS 'Наличие ограждений';
COMMENT ON COLUMN buildings_land_info.benches IS 'Количество скамеек';

CREATE TABLE buildings_history (
       term_id INTEGER NOT NULL REFERENCES terms ON DELETE CASCADE ON UPDATE CASCADE,
       LIKE buildings INCLUDING COMMENTS,
       copy_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
COMMENT ON COLUMN buildings_history.term_id IS 'Код месяца';
COMMENT ON COLUMN buildings_history.copy_time IS 'Время копирования';
ALTER TABLE buildings_history RENAME COLUMN id TO bldn_id;
ALTER TABLE buildings_history ADD CONSTRAINT buildings_history_bldn_id_fkey FOREIGN KEY (bldn_id) REFERENCES buildings(id) ON DELETE CASCADE ON UPDATE CASCADE;

CREATE TABLE buildings_tech_info_history (
       term_id INTEGER NOT NULL REFERENCES terms ON DELETE CASCADE ON UPDATE CASCADE,
       LIKE buildings_tech_info INCLUDING COMMENTS,
       copy_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
COMMENT ON COLUMN buildings_tech_info_history.term_id IS 'Код месяца';
COMMENT ON COLUMN buildings_tech_info_history.copy_time IS 'Время копирования';

CREATE TABLE buildings_land_info_history (
       term_id INTEGER NOT NULL REFERENCES terms ON DELETE CASCADE ON UPDATE CASCADE,
       LIKE buildings_land_info INCLUDING COMMENTS,
       copy_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
COMMENT ON COLUMN buildings_land_info_history.term_id IS 'Код месяца';
COMMENT ON COLUMN buildings_land_info_history.copy_time IS 'Время копирования';

CREATE TABLE uk_services (
  id SERIAL PRIMARY KEY,
  name VARCHAR(100)
);
COMMENT ON TABLE uk_services IS 'Услуги УК, по которым происходит начисление/сбор';
COMMENT ON COLUMN uk_services.id IS 'Код услуги';
COMMENT ON COLUMN uk_services.name IS 'Название услуги';

INSERT INTO uk_services(name) VALUES ('Содержание'), ('Взнос КР'), ('ХВ СОИ'), ('ГВ СОИ'), ('ЭЭ СОИ'), ('Прочие долги');

CREATE TABLE rkc_services (
  id SERIAL PRIMARY KEY,
  name VARCHAR(20),
  uk_service_id INTEGER REFERENCES uk_services ON DELETE CASCADE ON UPDATE CASCADE,
  full_name VARCHAR(100),
  total_uk_service_id INTEGER REFERENCES uk_services ON DELETE SET NULL ON UPDATE CASCADE
);
COMMENT ON TABLE rkc_services IS 'Сопоставление услуг УК и РКЦ';
COMMENT ON COLUMN rkc_services.id IS 'Код услуги РКЦ';
COMMENT ON COLUMN rkc_services.name IS 'Название услуги в РКЦ';
COMMENT ON COLUMN rkc_services.uk_service_id IS 'Код сопоставленной услуги УК';
COMMENT ON COLUMN rkc_services.full_name IS 'Полное название услуги в РКЦ';
COMMENT ON COLUMN rkc_services.total_uk_service_id IS 'Код сопоставленной услуги УК для отчета по собираемости';

CREATE TABLE uk_accrued_source (
  id SERIAL PRIMARY KEY,
  name VARCHAR
);

COMMENT ON TABLE uk_accrued_source IS 'Справочник организаций, высталяющих начисления';
COMMENT ON COLUMN uk_accrued_source.id IS 'Код организации';
COMMENT ON COLUMN uk_accrued_source.name IS 'Название';

INSERT INTO uk_accrued_source VALUES (1, 'РКЦ'), (2, 'Бухгалтерия'), (3, 'МО взнос КР'), (4, 'МО превышение');

CREATE TABLE rkc_values_history (
  bldn_id INTEGER REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE,
  term_id INTEGER REFERENCES terms ON DELETE CASCADE ON UPDATE CASCADE,
  rkc_service_id INTEGER REFERENCES rkc_services ON DELETE RESTRICT ON UPDATE CASCADE,
  occ_id INTEGER,
  flat_no VARCHAR(50),
  accrued NUMERIC(12,2),
  added NUMERIC(12,2),
  compens NUMERIC(12,2),
  paid NUMERIC(12,2),
  uk_accrued_source_id INTEGER DEFAULT 1 REFERENCES uk_accrued_source ON DELETE RESTRICT ON UPDATE CASCADE,
  in_saldo NUMERIC(12,2),
  out_saldo NUMERIC(12,2)
);
COMMENT ON TABLE rkc_values_history IS 'История начислений РКЦ';
COMMENT ON COLUMN rkc_values_history.bldn_id IS 'Код дома';
COMMENT ON COLUMN rkc_values_history.term_id IS 'Месяц';
COMMENT ON COLUMN rkc_values_history.rkc_service_id IS 'Услуга';
COMMENT ON COLUMN rkc_values_history.occ_id IS 'Лицевой счет';
COMMENT ON COLUMN rkc_values_history.flat_no IS 'Квартира';
COMMENT ON COLUMN rkc_values_history.accrued IS 'Начислено 100%';
COMMENT ON COLUMN rkc_values_history.added IS 'Перерасчеты';
COMMENT ON COLUMN rkc_values_history.compens IS 'Субсидия';
COMMENT ON COLUMN rkc_values_history.paid IS 'Оплата';
COMMENT ON COLUMN rkc_values_history.uk_accrued_source_id IS 'Кто выставил начисления';
COMMENT ON COLUMN rkc_values_history.in_saldo IS 'Входящее сальдо';
COMMENT ON COLUMN rkc_values_history.out_saldo IS 'Исходящее сальдо';

CREATE TABLE expense_groups (
  id INTEGER PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY
  , name VARCHAR(300) NOT NULL
  , report_priority INTEGER
  , parent_group INTEGER REFERENCES expense_groups ON DELETE SET NULL ON UPDATE CASCADE
);
COMMENT ON TABLE expense_groups IS 'Группы структуры платы';
COMMENT ON COLUMN expense_groups.id IS 'Код';
COMMENT ON COLUMN expense_groups.name IS 'Название группы';
COMMENT ON COLUMN expense_groups.report_priority IS 'Приоритет вывода в отчет';
COMMENT ON COLUMN expense_groups.parent_group IS 'Родительская группа';

CREATE TABLE expense_items (
  id INTEGER GENERATED BY DEFAULT AS IDENTITY,
  name1 VARCHAR(200),
  name2 VARCHAR(200),
  short_name VARCHAR(100),
  gis_guid VARCHAR(50),
  uk_service_id INTEGER NOT NULL REFERENCES uk_services ON DELETE CASCADE ON UPDATE CASCADE,
  group_id INTEGER NOT NULL REFERENCES expense_groups ON DELETE CASCADE ON UPDATE CASCADE,
  report_priority INTEGER DEFAULT 0,
  use_as_group_name BOOLEAN DEFAULT FALSE  
);
COMMENT ON TABLE expense_items IS 'Статьи расходов структуры платы';
COMMENT ON COLUMN expense_items.id IS 'Код';
COMMENT ON COLUMN expense_items.name1 IS 'Название1';
COMMENT ON COLUMN expense_items.name2 IS 'Название2';
COMMENT ON COLUMN expense_items.short_name IS 'Короткое название';
COMMENT ON COLUMN expense_items.gis_guid IS 'GUID в ГИС ЖКХ';
COMMENT ON COLUMN expense_items.uk_service_id IS 'Код услуги УК';
COMMENT ON COLUMN expense_items.group_id IS 'Код группы структуры платы';
COMMENT ON COLUMN expense_items.report_priority IS 'Приоритет вывода в отчет';
COMMENT ON COLUMN expense_items.use_as_group_name IS 'Флаг использования как название группы в отчете';
CREATE UNIQUE INDEX use_as_group_name_must_be_unique_on_group ON expense_items(group_id, use_as_group_name) WHERE use_as_group_name = true;

CREATE TABLE accrueds
       (bldn_id INTEGER REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE,
       contractor_id INTEGER REFERENCES contractors ON DELETE RESTRICT ON UPDATE CASCADE,
       mc_id INTEGER REFERENCES management_companies ON DELETE RESTRICT ON UPDATE CASCADE,
       acc_date INTEGER REFERENCES terms ON DELETE RESTRICT ON UPDATE CASCADE,
       acc_sum NUMERIC(12,2),
       CONSTRAINT bldn_acc_must_be_unique UNIQUE(bldn_id, acc_date));

CREATE TABLE work_financing_sources
       (id SERIAL PRIMARY KEY,
       name VARCHAR(100) NOT NULL,
       note TEXT,
       from_subaccount BOOLEAN DEFAULT FALSE
);

INSERT INTO work_financing_sources VALUES (0, 'Содержание',''), (1, 'Субсчет','');

SELECT SETVAL('work_financing_sources_id_seq', max(id)) FROM work_financing_sources;

CREATE TABLE works
       (id SERIAL PRIMARY KEY,
       gwt_id INTEGER REFERENCES global_work_types ON DELETE RESTRICT ON UPDATE CASCADE,
       workkind_id INTEGER REFERENCES work_kinds ON DELETE RESTRICT ON UPDATE CASCADE,
       bldn_id INTEGER REFERENCES buildings ON DELETE RESTRICT ON UPDATE CASCADE,
       work_date INTEGER REFERENCES terms ON DELETE RESTRICT ON UPDATE CASCADE,
       work_sum NUMERIC(12,2),
       si VARCHAR(40),
       volume VARCHAR(50),
       note TEXT,
       private_note TEXT,
       contractor_id INTEGER REFERENCES contractors ON DELETE RESTRICT ON UPDATE CASCADE,
       mc_id INTEGER REFERENCES management_companies ON DELETE RESTRICT ON UPDATE CASCADE,
       dogovor VARCHAR(200),
       finance_source INTEGER DEFAULT(1) REFERENCES work_financing_sources ON DELETE RESTRICT ON UPDATE RESTRICT,
       print_flag BOOLEAN DEFAULT TRUE,
       add_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
       change_date TIMESTAMP);

CREATE TABLE hidden_maintenance_works (
       id SERIAL PRIMARY KEY,
       man_hours NUMERIC(5,2),
       man_hour_mode_id INTEGER NOT NULL REFERENCES man_hour_cost_modes ON DELETE RESTRICT ON UPDATE CASCADE,
       workref_id INTEGER NOT NULL REFERENCES works ON DELETE CASCADE ON UPDATE CASCADE
);
COMMENT ON TABLE hidden_maintenance_works IS 'Работы по содержанию в разрезе материалов, транспорта и человекочасов';
COMMENT ON COLUMN hidden_maintenance_works.id IS 'Код работы';
COMMENT ON COLUMN hidden_maintenance_works.man_hours IS 'Количество человекочасов';
COMMENT ON COLUMN hidden_maintenance_works.man_hour_mode_id IS 'Код режима стоимости человекочаса';
COMMENT ON COLUMN hidden_maintenance_works.workref_id IS 'Ссылка на работу в таблице works';

CREATE TABLE plan_work_statuses
       (id INTEGER PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY,
       name VARCHAR(20) UNIQUE,
       in_plan BOOLEAN NOT NULL,
       can_new BOOLEAN NOT NULL,
       is_done BOOLEAN DEFAULT FALSE);
COMMENT ON TABLE plan_work_statuses IS 'Статусы планируемых работ';
COMMENT ON COLUMN plan_work_statuses.id IS 'Код статуса';
COMMENT ON COLUMN plan_work_statuses.name IS 'Название';
COMMENT ON COLUMN plan_work_statuses.in_plan IS 'Флаг, что статус считается "В плане"';
COMMENT ON COLUMN plan_work_statuses.can_new IS 'Флаг, что статус может быть у новой работы';
COMMENT ON COLUMN plan_work_statuses.is_done IS 'Флаг, что статус у выполненной работы';

INSERT INTO plan_work_statuses VALUES (1, 'В плане', TRUE, TRUE, FALSE), (2, 'В работе', TRUE, FALSE, FALSE), (3, 'Выполнена', FALSE, FALSE, TRUE), (4, 'Отменена', FALSE, FALSE, FALSE), (5, 'В стадии заключения', TRUE, FALSE, FALSE), (6, 'Накопление', FALSE, TRUE, FALSE);

SELECT SETVAL('plan_work_statuses_id_seq', max(id)) FROM plan_work_statuses;

CREATE TABLE plan_works
       (id SERIAL PRIMARY KEY,
       gwt_id INTEGER REFERENCES global_work_types ON DELETE RESTRICT ON UPDATE CASCADE,
       workkind_id INTEGER REFERENCES work_kinds ON DELETE RESTRICT ON UPDATE CASCADE,
       bldn_id INTEGER REFERENCES buildings ON DELETE RESTRICT ON UPDATE CASCADE,
       work_date DATE,
       work_sum NUMERIC(10,2),
       smeta_sum NUMERIC(10,2),
       note TEXT,
       private_note TEXT,
       contractor_id INTEGER REFERENCES contractors ON DELETE RESTRICT ON UPDATE CASCADE,
       mc_id INTEGER REFERENCES management_companies ON DELETE RESTRICT ON UPDATE CASCADE,
       work_status INTEGER REFERENCES plan_work_statuses ON DELETE RESTRICT ON UPDATE CASCADE,
       employee VARCHAR(200),
       begin_date DATE DEFAULT (NULL),
       end_date DATE DEFAULT (NULL),
       work_ref INTEGER DEFAULT (NULL) REFERENCES works(id) ON DELETE SET NULL ON UPDATE CASCADE,
       create_user INTEGER REFERENCES users(id) ON DELETE RESTRICT ON UPDATE CASCADE,
       last_change_user INTEGER REFERENCES users ON DELETE RESTRICT ON UPDATE CASCADE
);
COMMENT ON TABLE plan_works IS 'Планируемые работы';
COMMENT ON COLUMN plan_works.id IS 'Код работы';
COMMENT ON COLUMN plan_works.gwt_id IS 'Тип ремонта';
COMMENT ON COLUMN plan_works.workkind_id IS 'Вид работы';
COMMENT ON COLUMN plan_works.bldn_id IS 'Код дома';
COMMENT ON COLUMN plan_works.work_date IS 'Планируемый месяц';
COMMENT ON COLUMN plan_works.work_sum IS 'Сумма работы';
COMMENT ON COLUMN plan_works.smeta_sum IS 'Сумма работы по смете';
COMMENT ON COLUMN plan_works.note IS 'Комментарий';
COMMENT ON COLUMN plan_works.contractor_id IS 'Код подрядчика';
COMMENT ON COLUMN plan_works.mc_id IS 'Код УК';
COMMENT ON COLUMN plan_works.work_status IS 'Статус работы';
COMMENT ON COLUMN plan_works.employee IS 'Ответственный работник';
COMMENT ON COLUMN plan_works.begin_date IS 'Дата начала работы';
COMMENT ON COLUMN plan_works.end_date IS 'Дата окончания работы';
COMMENT ON COLUMN plan_works.work_ref IS 'Код выполненной работы';
COMMENT ON COLUMN plan_works.create_user IS 'Пользователь, создавший работу';
COMMENT ON COLUMN plan_works.last_change_user IS 'Последний кто менял работу';
COMMENT ON COLUMN plan_works.private_note IS 'Внутренний комментарий';


CREATE TABLE old_works
       (id SERIAL PRIMARY KEY,
       bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE RESTRICT ON UPDATE CASCADE,
       work_name VARCHAR(200) NOT NULL,
       work_year INTEGER NOT NULL,
       work_volume VARCHAR(100) NOT NULL,
       work_sum NUMERIC(10,2),
       note TEXT,
       other_budget_flag INTEGER NOT NULL DEFAULT (0),
       other_budget_note VARCHAR(30));

CREATE TABLE expenses
       (id SERIAL PRIMARY KEY,
       expense_item INTEGER NOT NULL REFERENCES expense_items ON DELETE RESTRICT ON UPDATE CASCADE,
       term_id INTEGER NOT NULL REFERENCES terms ON DELETE RESTRICT ON UPDATE CASCADE,
       bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE,
       price NUMERIC(6, 2),
       expense_plan_sum NUMERIC(14, 2),
       expense_fact_sum NUMERIC(14, 2),
       CONSTRAINT expense_in_term_must_be_unique UNIQUE (expense_item, term_id, bldn_id));

CREATE TABLE bldn_expense_names
       (expense_item INTEGER NOT NULL REFERENCES expense_items ON DELETE CASCADE ON UPDATE CASCADE,
       bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE,
       name_use INTEGER NOT NULL DEFAULT(1),
       CONSTRAINT bldn_expense_name_must_be_unique UNIQUE (expense_item, bldn_id));

CREATE TABLE bldn_services
       (bldn_id INTEGER REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE,
       service_id INTEGER REFERENCES services ON DELETE RESTRICT ON UPDATE CASCADE,
       mode_id INTEGER REFERENCES service_modes ON DELETE RESTRICT ON UPDATE CASCADE,
       inputs_count INTEGER DEFAULT 0,
       possible_counter BOOLEAN DEFAULT FALSE,
       note TEXT,
       CONSTRAINT bldn_service_must_be_unique UNIQUE (bldn_id, service_id));
COMMENT ON TABLE bldn_services IS 'Режимы ЖКУ в доме';
COMMENT ON COLUMN bldn_services.bldn_id IS 'Код дома';
COMMENT ON COLUMN bldn_services.service_id IS 'Код услуги';
COMMENT ON COLUMN bldn_services.mode_id IS 'Код режима';
COMMENT ON COLUMN bldn_services.inputs_count IS 'Количество вводов';
COMMENT ON COLUMN bldn_services.possible_counter IS 'Возможность установки прибора учета';
COMMENT ON COLUMN bldn_services.note IS 'Примечание';

CREATE TABLE bldn_services_history (
  term_id INTEGER NOT NULL REFERENCES terms(id) ON UPDATE CASCADE ON DELETE CASCADE
  , bldn_id INTEGER NOT NULL REFERENCES buildings(id) ON UPDATE CASCADE ON DELETE CASCADE
  , service_id INTEGER NOT NULL REFERENCES services(id) ON UPDATE CASCADE ON DELETE CASCADE
  , mode_id INTEGER NOT NULL REFERENCES service_modes(id) ON UPDATE CASCADE ON DELETE CASCADE
  , inputs_count INTEGER
  , possible_counter BOOLEAN
  , note TEXT
  , CONSTRAINT bldn_service_term_must_be_unique UNIQUE (term_id, bldn_id, service_id)
);
COMMENT ON TABLE bldn_services_history IS 'История режимов ЖКУ в доме';
COMMENT ON COLUMN bldn_services_history.term_id IS 'Код периода';
COMMENT ON COLUMN bldn_services_history.bldn_id IS 'Код дома';
COMMENT ON COLUMN bldn_services_history.service_id IS 'Код услуги';
COMMENT ON COLUMN bldn_services_history.mode_id IS 'Код режима';
COMMENT ON COLUMN bldn_services_history.inputs_count IS 'Количество вводов';
COMMENT ON COLUMN bldn_services_history.possible_counter IS 'Возможность установки прибора учета';
COMMENT ON COLUMN bldn_services_history.note IS 'Примечание';

CREATE TABLE bldn_subaccounts
       (bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE,
       term_id INTEGER NOT NULL REFERENCES terms ON DELETE RESTRICT ON UPDATE CASCADE,
       subaccount_sum NUMERIC(14, 2),
       CONSTRAINT bldn_subaccount_unique UNIQUE(bldn_id, term_id));

CREATE TABLE plan_subaccounts (
       bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE CONSTRAINT bldn_plan_subaccount_must_be_unique UNIQUE,
       plan_sum NUMERIC(12, 2),
       plan_percent NUMERIC(3, 2)
);

CREATE TABLE counter_models (
       id SERIAL PRIMARY KEY,
       model_name VARCHAR(300),
       has_dti BOOLEAN DEFAULT FALSE,
       calibration_interval INTEGER);

CREATE TABLE work_material_types (
       id SERIAL PRIMARY KEY,
       material_name VARCHAR NOT NULL CONSTRAINT material_name_must_be_unique UNIQUE,
       is_transport BOOLEAN NOT NULL DEFAULT FALSE
);
COMMENT ON TABLE work_material_types IS 'Материалы для работ';
COMMENT ON COLUMN work_material_types.id IS 'Код материала';
COMMENT ON COLUMN work_material_types.material_name IS 'Наименование материала';
COMMENT ON COLUMN work_material_types.is_transport IS 'Признак транспорта';
INSERT INTO work_material_types VALUES (0, 'Не указано', FALSE), (1, 'Транспорт', TRUE);
SELECT SETVAL('work_material_types_id_seq', MAX(id)) FROM work_material_types;
CREATE UNIQUE INDEX lower_work_material_types_must_be_unique_idx ON work_material_types (LOWER(material_name));

CREATE TABLE works_materials (
       id SERIAL PRIMARY KEY,
       maintenance_work_id INTEGER REFERENCES hidden_maintenance_works ON DELETE CASCADE ON UPDATE CASCADE,
       material_id INTEGER REFERENCES work_material_types ON DELETE RESTRICT ON UPDATE CASCADE DEFERRABLE,
       material_note TEXT,
       material_cost NUMERIC(8, 2),
       material_count NUMERIC(5, 2),
       material_si TEXT
);
COMMENT ON TABLE works_materials IS 'Материалы работ';
COMMENT ON COLUMN works_materials.id IS 'Код';
COMMENT ON COLUMN works_materials.maintenance_work_id IS 'Код работы по содержанию';
COMMENT ON COLUMN works_materials.material_id IS 'Код материала';
COMMENT ON COLUMN works_materials.material_note IS 'Примечание к материалу';
COMMENT ON COLUMN works_materials.material_cost IS 'Цена за единицу материала';
COMMENT ON COLUMN works_materials.material_count IS 'Количество материалов';
COMMENT ON COLUMN works_materials.material_si IS 'Единица изменения материалов';

CREATE TABLE tmp_counters (
  id SERIAL PRIMARY KEY,
  bldn_id INTEGER REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE,
  name VARCHAR
);
COMMENT ON TABLE tmp_counters IS 'Приборы учёта';
COMMENT ON COLUMN tmp_counters.id IS 'Код';
COMMENT ON COLUMN tmp_counters.bldn_id IS 'Код здания';
COMMENT ON COLUMN tmp_counters.name IS 'Название';

CREATE TABLE certificates (
  id SERIAL PRIMARY KEY,
  counter_id INTEGER REFERENCES tmp_counters ON DELETE CASCADE ON UPDATE CASCADE,
  certificate_date DATE,
  certificate_validite INTEGER DEFAULT 0,
  note TEXT
);
COMMENT ON TABLE certificates IS 'Акты допуска';
COMMENT ON COLUMN certificates.id IS 'Код';
COMMENT ON COLUMN certificates.counter_id IS 'Код прибора учёта';
COMMENT ON COLUMN certificates.certificate_date IS 'Дата акта';
COMMENT ON COLUMN certificates.certificate_validite IS 'Срок действия акта';
COMMENT ON COLUMN certificates.note IS 'Примечание';

CREATE TABLE bldn_id_mapping (
  bldn_id INTEGER REFERENCES buildings(id) ON DELETE CASCADE ON UPDATE CASCADE,
  energosbyt_bldn_id INTEGER,
  UNIQUE (bldn_id, energosbyt_bldn_id)
);

COMMENT ON TABLE bldn_id_mapping IS 'Сопоставление кода дома с внешними кодами';
COMMENT ON COLUMN bldn_id_mapping.bldn_id IS 'Наш код дома';
COMMENT ON COLUMN bldn_id_mapping.energosbyt_bldn_id IS 'Код энергосбыт Волга';


CREATE TABLE meter_readings (
  bldn_id INTEGER REFERENCES buildings(id) ON DELETE CASCADE ON UPDATE CASCADE,
  flat_no VARCHAR(30),
  service_id INTEGER REFERENCES services(id) ON DELETE RESTRICT ON UPDATE CASCADE,
  term_id INTEGER REFERENCES terms(id) ON DELETE CASCADE ON UPDATE CASCADE,
  readings REAL,
  UNIQUE(bldn_id, flat_no, term_id, service_id)
);
COMMENT ON TABLE meter_readings IS 'Показания приборов учёта';
COMMENT ON COLUMN meter_readings.bldn_id IS 'Код дома';
COMMENT ON COLUMN meter_readings.flat_no IS 'Номер квартиры';
COMMENT ON COLUMN meter_readings.service_id IS 'Код услуги';
COMMENT ON COLUMN meter_readings.term_id IS 'Код периода';
COMMENT ON COLUMN meter_readings.readings IS 'Показания';

CREATE TABLE meter_readings_tmp (
  bldn_id INTEGER,
  flat_no VARCHAR(30),
  term_id INTEGER,
  readings REAL,
  service_id INTEGER
);
COMMENT ON TABLE meter_readings_tmp IS 'Временная таблица для загрузки покзаний из файла';

CREATE TABLE flats (
  flat_id BIGINT GENERATED BY DEFAULT AS IDENTITY
  , term_id INTEGER REFERENCES terms ON DELETE CASCADE ON UPDATE CASCADE
  , bldn_id INTEGER REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE
  , flat_no VARCHAR(100)
  , residental BOOLEAN DEFAULT TRUE
  , uninhabitable BOOLEAN DEFAULT FALSE
  , rooms INTEGER
  , passport_square NUMERIC(6, 2)
  , square NUMERIC(6, 2)
  , note TEXT
  , cadastral_no TEXT
  , PRIMARY KEY (flat_id, term_id)
  , CONSTRAINT flat_number_in_term_must_be_unique UNIQUE(bldn_id, flat_no, term_id)
);
COMMENT ON TABLE flats IS 'Помещения';
COMMENT ON COLUMN flats.flat_id IS 'Код квартиры в периоде';
COMMENT ON COLUMN flats.term_id IS 'Период';
COMMENT ON COLUMN flats.bldn_id IS 'Код дома';
COMMENT ON COLUMN flats.flat_no IS 'Номер квартиры';
COMMENT ON COLUMN flats.residental IS 'Жилое/Нежилое';
COMMENT ON COLUMN flats.uninhabitable IS 'Непригодно для проживания';
COMMENT ON COLUMN flats.rooms IS 'Комнат';
COMMENT ON COLUMN flats.passport_square IS 'Площадь по тех. паспорту';
COMMENT ON COLUMN flats.square IS 'Площадь';
COMMENT ON COLUMN flats.note IS 'Примечание';
COMMENT ON COLUMN flats.cadastral_no IS 'Кадастровый номер';

CREATE INDEX IF NOT EXISTS flats_term_bldn_idx ON flats(term_id, bldn_id);

CREATE TABLE flat_shares (
  id BIGINT PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY
  ,term_id INTEGER
  ,flat_id BIGINT
  ,share_numerator INTEGER
  ,share_denominator INTEGER
  ,is_legal_entity BOOLEAN DEFAULT FALSE
  ,is_privatized BOOLEAN
  ,in_period_id INTEGER NOT NULL
  ,FOREIGN KEY (term_id, flat_id) REFERENCES flats(term_id, flat_id) ON DELETE CASCADE ON UPDATE CASCADE
);
COMMENT ON TABLE flat_shares IS 'Доли квартир по документам';
COMMENT ON COLUMN flat_shares.id IS 'Код для привязки собственников';
COMMENT ON COLUMN flat_shares.term_id IS 'Период';
COMMENT ON COLUMN flat_shares.flat_id IS 'Код квартиры';
COMMENT ON COLUMN flat_shares.share_numerator IS 'Числитель доли';
COMMENT ON COLUMN flat_shares.share_denominator IS 'Знаменатель доли';
COMMENT ON COLUMN flat_shares.is_legal_entity IS 'Юр.лицо';
COMMENT ON COLUMN flat_shares.is_privatized IS 'Приватизация';
CREATE UNIQUE INDEX flat_shares_term_in_period_index ON flat_shares (term_id, in_period_id);

CREATE TABLE owners (
  id BIGINT PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY
  , share_id BIGINT REFERENCES flat_shares ON DELETE CASCADE ON UPDATE CASCADE
  , owner_name VARCHAR(300)
  , owner_document VARCHAR(300)
  , phone VARCHAR(300)
  , has_pd_consent BOOLEAN
  , is_chairman BOOLEAN DEFAULT FALSE
  , is_sekretar BOOLEAN DEFAULT FALSE
  , is_senat BOOLEAN DEFAULT FALSE
);
COMMENT ON TABLE owners IS 'Собственники';
COMMENT ON COLUMN owners.id IS 'Код';
COMMENT ON COLUMN owners.share_id IS 'Код доли';
COMMENT ON COLUMN owners.owner_name IS 'Имя собственника';
COMMENT ON COLUMN owners.owner_document IS 'Правоустанавливающий документ';
COMMENT ON COLUMN owners.phone IS 'Телефон';
COMMENT ON COLUMN owners.has_pd_consent IS 'Наличие согласия на обработку ПД';
COMMENT ON COLUMN owners.is_chairman IS 'Комендант';
COMMENT ON COLUMN owners.is_sekretar IS 'Секретарь';
COMMENT ON COLUMN owners.is_senat IS 'Совет дома';
CREATE INDEX IF NOT EXISTS owners_share_id_idx ON owners(share_id) WITH (deduplicate_items=OFF);

CREATE TABLE rkc_added_types (
  id INTEGER PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY
  , name VARCHAR(100)
);
COMMENT ON TABLE rkc_added_types IS 'Типы разовых начислений';
COMMENT ON COLUMN rkc_added_types.id IS 'код';
COMMENT ON COLUMN rkc_added_types.name IS 'название';
INSERT INTO rkc_added_types VALUES (1, 'Снятия комендантам'), (2, 'Перерасчет за уборку'), (3, 'Списание задолженности');

CREATE TABLE rkc_addeds_history (
  type_id INTEGER REFERENCES rkc_added_types ON DELETE CASCADE ON UPDATE CASCADE
  , term_id INTEGER NOT NULL
  , occ_id INTEGER NOT NULL
  , bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE
  , service_id INTEGER NOT NULL REFERENCES rkc_services ON DELETE RESTRICT ON UPDATE CASCADE
  , added_value NUMERIC(12, 2)
);
COMMENT ON TABLE rkc_addeds_history IS 'Разовые начисления';
COMMENT ON COLUMN rkc_addeds_history.type_id IS 'тип разового';
COMMENT ON COLUMN rkc_addeds_history.term_id IS 'период';
COMMENT ON COLUMN rkc_addeds_history.occ_id IS 'лицевой счет';
COMMENT ON COLUMN rkc_addeds_history.service_id IS 'услуга';
COMMENT ON COLUMN rkc_addeds_history.added_value IS 'сумма';

CREATE TABLE common_property_group (
  id INTEGER PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY
  , name VARCHAR(100) NOT NULL
);
COMMENT ON TABLE common_property_group IS 'Группы элементов общего имущества';
COMMENT ON COLUMN common_property_group.id IS 'код';
COMMENT ON COLUMN common_property_group.name IS 'название группы';

CREATE TABLE common_property_element (
  id INTEGER PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY
  , group_id INTEGER NOT NULL REFERENCES common_property_group ON DELETE RESTRICT ON UPDATE CASCADE
  , name VARCHAR(200) NOT NULL
  , is_required BOOLEAND DEFAULT TRUE
  , CONSTRAINT element_name_must_be_unique UNIQUE (group_id, name)
);
COMMENT ON TABLE common_property_element IS 'Элементы общего имущества дома';
COMMENT ON COLUMN common_property_element.id IS 'код элемента';
COMMENT ON COLUMN common_property_element.group_id IS 'код группы элемента';
COMMENT ON COLUMN common_property_element.name IS 'название элемента';
COMMENT ON COLUMN common_property_element.is_required IS 'Обязательный параметр';

CREATE TABLE building_common_property_elements (
  bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE
  , element_id INTEGER NOT NULL REFERENCES common_property_element ON DELETE CASCADE ON UPDATE CASCADE
  , is_contain BOOLEAN DEFAULT FALSE
  , element_state TEXT
  , PRIMARY KEY (bldn_id, element_id)
);
COMMENT ON TABLE building_common_property_elements IS 'Состояние элементов общего имущества дома';
COMMENT ON COLUMN building_common_property_elements.bldn_id IS 'код дома';
COMMENT ON COLUMN building_common_property_elements.element_id IS 'код элемента';
COMMENT ON COLUMN building_common_property_elements.is_contain IS 'имеется ли элемент в доме';
COMMENT ON COLUMN building_common_property_elements.element_state IS 'состояние элемента';

CREATE TABLE building_common_property_elements_history (
  term_id INTEGER NOT NULL REFERENCES terms ON DELETE CASCADE ON UPDATE CASCADE
  , bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE
  , element_id INTEGER NOT NULL REFERENCES common_property_element ON DELETE CASCADE ON UPDATE CASCADE
  , is_contain BOOLEAN DEFAULT FALSE
  , element_state TEXT
  , PRIMARY KEY (term_id, bldn_id, element_id)
);
COMMENT ON TABLE building_common_property_elements_history IS 'История состояний элементов общего имущества дома';
COMMENT ON COLUMN building_common_property_elements_history.term_id IS 'код периода';
COMMENT ON COLUMN building_common_property_elements_history.bldn_id IS 'код дома';
COMMENT ON COLUMN building_common_property_elements_history.element_id IS 'код элемента';
COMMENT ON COLUMN building_common_property_elements_history.is_contain IS 'имеется ли элемент в доме';
COMMENT ON COLUMN building_common_property_elements_history.element_state IS 'состояние элемента';

CREATE TABLE common_property_element_parameter (
  id BIGINT PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY
  , element_id INTEGER NOT NULL REFERENCES common_property_element ON DELETE CASCADE ON UPDATE CASCADE
  , name VARCHAR(100) NOT NULL
  , CONSTRAINT parameter_name_must_be_unique UNIQUE (element_id, name)
);
COMMENT ON TABLE common_property_element_parameter IS 'Параметры элементов общего имущества дома';
COMMENT ON COLUMN common_property_element_parameter.id IS 'код параметра';
COMMENT ON COLUMN common_property_element_parameter.element_id IS 'код элемента';
COMMENT ON COLUMN common_property_element_parameter.name IS 'название параметра';

CREATE TABLE building_common_property_element_parameter (
  bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE
  , parameter_id BIGINT NOT NULL REFERENCES common_property_element_parameter ON DELETE CASCADE ON UPDATE CASCADE
  , parameter_value TEXT
  , is_using BOOLEAN DEFAULT TRUE
  , PRIMARY KEY (bldn_id, parameter_id)
);
COMMENT ON TABLE building_common_property_element_parameter IS 'Значения параметров элементов общего имущества дома';
COMMENT ON COLUMN building_common_property_element_parameter.bldn_id IS 'код дома';
COMMENT ON COLUMN building_common_property_element_parameter.parameter_id IS 'код параметра';
COMMENT ON COLUMN building_common_property_element_parameter.parameter_value IS 'значение параметра';

CREATE TABLE building_common_property_element_parameter_history (
  term_id INTEGER NOT NULL REFERENCES terms ON DELETE CASCADE ON UPDATE CASCADE
  , bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE
  , parameter_id BIGINT NOT NULL REFERENCES common_property_element_parameter ON DELETE CASCADE ON UPDATE CASCADE
  , parameter_value TEXT
  , PRIMARY KEY (term_id, bldn_id, parameter_id)
);
COMMENT ON TABLE building_common_property_element_parameter_history IS 'История значений параметров элементов общего имущества дома';
COMMENT ON COLUMN building_common_property_element_parameter_history.term_id IS 'код периода';
COMMENT ON COLUMN building_common_property_element_parameter_history.bldn_id IS 'код дома';
COMMENT ON COLUMN building_common_property_element_parameter_history.parameter_id IS 'код параметра';
COMMENT ON COLUMN building_common_property_element_parameter_history.parameter_value IS 'значение параметра';

CREATE TABLE log_offers (
  bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE
  , action_id INTEGER NOT NULL REFERENCES log_operations ON DELETE RESTRICT ON UPDATE CASCADE
  , action_time TIMESTAMP DEFAULT LOCALTIMESTAMP
  , user_id INTEGER NOT NULL REFERENCES users ON DELETE RESTRICT ON UPDATE CASCADE
  , pc_name VARCHAR(100) NOT NULL
  , action_description TEXT
);
COMMENT ON TABLE log_offers IS 'История изменений предложений';
COMMENT ON COLUMN log_offers.bldn_id IS 'Код дома';
COMMENT ON COLUMN log_offers.action_id IS 'Код операции';
COMMENT ON COLUMN log_offers.action_time IS 'Время операциии';
COMMENT ON COLUMN log_offers.user_id IS 'Код пользователя';
COMMENT ON COLUMN log_offers.pc_name IS 'Имя компьютера';
COMMENT ON COLUMN log_offers.action_description IS 'Описание операции';

CREATE TABLE offers_work (
  id BIGINT PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY
  , offers_year DATE
  , bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE
  , work_name VARCHAR(300) NOT NULL
  , work_sum NUMERIC(12,2)
  , priority INTEGER
);
COMMENT ON TABLE offers_work IS 'Предлагаемые работы';
COMMENT ON COLUMN offers_work.id IS 'Код';
COMMENT ON COLUMN offers_work.offers_year IS 'На какой год предлагаем';
COMMENT ON COLUMN offers_work.bldn_id IS 'Дом';
COMMENT ON COLUMN offers_work.work_name IS 'Название работы';
COMMENT ON COLUMN offers_work.work_sum IS 'Сумма';
COMMENT ON COLUMN offers_work.priority IS 'Приоритет';

CREATE TABLE offers_expense (
  bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE
  , offers_date DATE
  , expense_item INTEGER NOT NULL REFERENCES expense_items ON DELETE CASCADE ON UPDATE CASCADE
  , expense_value NUMERIC(10,2)
);
COMMENT ON TABLE offers_expense IS 'Предлагаемая структура';
COMMENT ON COLUMN offers_expense.bldn_id IS 'Дом';
COMMENT ON COLUMN offers_expense.offers_date IS 'На какой год предлагаем';
COMMENT ON COLUMN offers_expense.expense_item IS 'Статья расходов';
COMMENT ON COLUMN offers_expense.expense_value IS 'Значение';

CREATE TABLE offers_annex (
  id BIGINT PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY
  , bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE
  , offer_text XML
  , annex_date TIMESTAMP DEFAULT LOCALTIMESTAMP
  , user_id INTEGER NOT NULL REFERENCES users ON DELETE CASCADE ON UPDATE CASCADE
  , pc_name VARCHAR
);
COMMENT ON TABLE offers_annex IS 'Распечатанные предложения';
COMMENT ON COLUMN offers_annex.id IS 'Код';
COMMENT ON COLUMN offers_annex.bldn_id IS 'Дом';
COMMENT ON COLUMN offers_annex.offer_text IS 'xml предложения';
COMMENT ON COLUMN offers_annex.annex_date IS 'дата вывода';
COMMENT ON COLUMN offers_annex.user_id IS 'пользователь';

CREATE TABLE man_hour_cost_modes (
  id INTEGER PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY
  , name VARCHAR(50) NOT NULL UNIQUE
);
COMMENT ON TABLE man_hour_cost_modes IS 'Режимы стоимости человекочаса';
COMMENT ON COLUMN man_hour_cost_modes.id IS 'Код режима';
COMMENT ON COLUMN man_hour_cost_modes.name IS 'Название';
INSERT INTO man_hour_cost_modes(id, name) VALUES (0, 'Нет');

CREATE TABLE man_hour_cost_rates (
  id BIGINT PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY
  , mode_id INTEGER NOT NULL REFERENCES man_hour_cost_modes(id) ON DELETE CASCADE ON UPDATE CASCADE
  , term_id INTEGER NOT NULL REFERENCES terms(id) ON DELETE RESTRICT ON UPDATE CASCADE
  , contractor_id INTEGER NOT NULL REFERENCES contractors(id) ON DELETE CASCADE ON UPDATE CASCADE
  , cost_sum NUMERIC(8, 2) NOT NULL DEFAULT 0.00
  , CONSTRAINT man_hour_cost_in_term_must_be_unique UNIQUE (mode_id, term_id, contractor_id)
);
COMMENT ON TABLE man_hour_cost_rates IS 'Ставки стоимости человекочаса';
COMMENT ON COLUMN man_hour_cost_rates.id IS 'Код элемента';
COMMENT ON COLUMN man_hour_cost_rates.mode_id IS 'Код режима';
COMMENT ON COLUMN man_hour_cost_rates.term_id IS 'Код периода';
COMMENT ON COLUMN man_hour_cost_rates.contractor_id IS 'Код подрядной организации';
COMMENT ON COLUMN man_hour_cost_rates.cost_sum IS 'Стоимость человекочаса';

CREATE TABLE bldn_man_hour_cost (
  term_id INTEGER NOT NULL REFERENCES terms ON DELETE RESTRICT ON UPDATE CASCADE
  , bldn_id INTEGER NOT NULL REFERENCES buildings ON DELETE RESTRICT ON UPDATE CASCADE
  , mode_id INTEGER NOT NULL REFERENCES man_hour_cost_modes ON DELETE RESTRICT ON UPDATE CASCADE
  , CONSTRAINT man_hour_cost_in_bldn_in_term_must_be_unique UNIQUE (term_id, bldn_id)
);
COMMENT ON TABLE bldn_man_hour_cost IS 'Стоимость человекочаса в доме';
COMMENT ON COLUMN bldn_man_hour_cost.term_id IS 'Код периода';
COMMENT ON COLUMN bldn_man_hour_cost.bldn_id IS 'Код дома';
COMMENT ON COLUMN bldn_man_hour_cost.mode_id IS 'Код режима человекочаса';

CREATE TABLE lawsuits (
  id BIGINT PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY
  , occ_id INTEGER NOT NULL
  , address TEXT
  , begin_period DATE
  , end_period DATE
  , lawsuit_number VARCHAR(100)
  , lawsuit_sum NUMERIC(10, 2)
  , fee_sum NUMERIC(10, 2)
  , is_paided BOOLEAN NOT NULL DEFAULT FALSE
  , is_active BOOLEAN NOT NULL DEFAULT TRUE
  , note TEXT
);
COMMENT ON TABLE lawsuits IS 'Судебные приказы';
COMMENT ON COLUMN lawsuits.id IS 'код приказа';
COMMENT ON COLUMN lawsuits.occ_id IS 'лицевой счет';
COMMENT ON COLUMN lawsuits.address IS 'адрес';
COMMENT ON COLUMN lawsuits.begin_period IS 'период начала задолженности';
COMMENT ON COLUMN lawsuits.end_period IS 'период окончания задолженности';
COMMENT ON COLUMN lawsuits.lawsuit_number IS 'номер иска';
COMMENT ON COLUMN lawsuits.lawsuit_sum IS 'сумма иска';
COMMENT ON COLUMN lawsuits.fee_sum IS 'сумма госпошлины';
COMMENT ON COLUMN lawsuits.is_paided IS 'признак оплачен';
COMMENT ON COLUMN lawsuits.is_active IS 'признак активности';
COMMENT ON COLUMN lawsuits.note IS 'примечание к приказу';

CREATE TABLE lawsuit_persons (
  id BIGINT PRIMARY KEY GENERATED BY DEFAULT AS IDENTITY
  , lawsuit_id BIGINT REFERENCES lawsuits(id) ON DELETE CASCADE ON UPDATE CASCADE
  , last_name VARCHAR(100)
  , first_name VARCHAR(50)
  , second_name VARCHAR(50)
);
COMMENT ON TABLE lawsuit_persons IS 'Люди, на которых есть судебные приказы';
COMMENT ON COLUMN lawsuit_persons.id IS 'идентификатор';
COMMENT ON COLUMN lawsuit_persons.lawsuit_id IS 'идентификатор приказа';
COMMENT ON COLUMN lawsuit_persons.last_name IS 'фамилия';
COMMENT ON COLUMN lawsuit_persons.first_name IS 'имя';
COMMENT ON COLUMN lawsuit_persons.second_name IS 'отчество';

CREATE TABLE chairman_signature (
  begin_term INTEGER REFERENCES terms ON DELETE RESTRICT ON UPDATE CASCADE,
  bldn_id INTEGER REFERENCES buildings ON DELETE CASCADE ON UPDATE CASCADE,
  sign BYTEA,
  signature_owner VARCHAR(100) NOT NULL,
  CONSTRAINT chairman_signature_term_bldn_unique UNIQUE (begin_term, bldn_id)
);
COMMENT ON TABLE chairman_signature IS 'Подписи комендантов';
COMMENT ON COLUMN chairman_signature.begin_term IS 'Код периода с которого действует';
COMMENT ON COLUMN chairman_signature.bldn_id IS 'Код дома';
COMMENT ON COLUMN chairman_signature.sign IS 'Картинка подписи';
COMMENT ON COLUMN chairman_signature.signature_owner IS 'ФИО владельца подписи';

-- BEGIN FUNCTIONS
CREATE PROCEDURE write_action_to_log(action_id INTEGER, user_id INTEGER, pc_name VARCHAR, log_action JSONB)
LANGUAGE SQL
AS $$
    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
		VALUES (action_id, user_id, pc_name, log_action);
$$;
COMMENT ON PROCEDURE write_action_to_log IS 'Добавление записи в лог';

CREATE FUNCTION get_error_number(InErrName VARCHAR) RETURNS VARCHAR AS
$$
  SELECT err_number FROM errors WHERE name = InErrName;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION get_error_message(InErrName VARCHAR) RETURNS VARCHAR AS
$$
  SELECT message FROM errors WHERE name = InErrName;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION is_all_values(InValue INTEGER) RETURNS BOOL AS
$$
  BEGIN
    RETURN InValue = CAST(c.value AS INTEGER)
      FROM constants c
     WHERE name = 'all_values';
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_not_value() RETURNS INTEGER AS
$$
  SELECT CAST (value AS INTEGER)
  FROM constants
  WHERE name = 'not_value';
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION is_not_value(InValue INTEGER) RETURNS BOOL AS
$$
  BEGIN
    RETURN InValue = CAST(c.value AS INTEGER)
      FROM constants c
     WHERE name = 'not_value';
  END;
$$ LANGUAGE plpgsql;


CREATE FUNCTION rights_get_counter_rights_number() RETURNS INTEGER AS
$$
  SELECT id FROM access_rights WHERE name = 'Приборы учёта';
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION rights_get_certificate_rights_number() RETURNS INTEGER AS
$$
  SELECT id FROM access_rights WHERE name = 'Акты ОДПУ';
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION rights_get_work_rights_number(GwtId INTEGER) RETURNS INTEGER AS
$$
  SELECT CAST(value AS INTEGER) + GwtId
  FROM constants
  WHERE name = 'gwt_work_access_prefix';
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION rights_get_work_rights_number IS 'Номер доступа к работам';

CREATE FUNCTION rights_get_plan_rights_number(GwtId INTEGER) RETURNS INTEGER AS
$$
  SELECT CAST(value AS INTEGER) + GwtId
  FROM constants
  WHERE name = 'gwt_planwork_access_prefix';
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION rights_get_bldn_accrued_rights_number() RETURNS INTEGER AS
$$
  SELECT id FROM access_rights WHERE name = 'Начисления услуг';
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION rights_get_contractor_cost_rights_number() RETURNS INTEGER AS
$$
  SELECT id FROM access_rights WHERE name = 'Расходы подрядчиков';
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION rights_get_meter_readings_rights_number() RETURNS INTEGER AS
$$
  SELECT id FROM access_rights WHERE name = 'Показания ИПУ';
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION rights_get_owners_rights_number() RETURNS INTEGER AS
$$
  SELECT id FROM access_rights WHERE name = 'Собственники помещений';
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION rights_get_common_property_elements_rights_number() RETURNS INTEGER AS
$$
  SELECT id FROM access_rights WHERE name = 'Элементы общего имущества';
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION rights_get_common_property_elements_rights_number IS 'Получение индекса прав на изменение справочников общего имущества';

CREATE FUNCTION rights_get_common_property_group_rights_number(GroupId INTEGER) RETURNS INTEGER AS
$$
  SELECT value::INTEGER + GroupId
  FROM constants
  WHERE name = 'common_property_group_prefix';
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION rights_get_common_property_group_rights_number IS 'Получение индекса прав на изменение элементов общего имущества в группе';

CREATE FUNCTION rights_get_tech_info_rights_number() RETURNS INTEGER AS
$$
  SELECT id FROM access_rights WHERE name = 'Техническая информация';
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION rights_get_tech_info_rights_number IS 'Получение индекса прав на техническую информацию';

CREATE FUNCTION get_file_version(InFileLoadType VARCHAR, OUT OutFileVersion INTEGER) AS
$$
  SELECT file_version FROM load_files_versions WHERE id = InFileLoadType;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_file_version IS 'Получение версии загружаемого файла с данными';


CREATE FUNCTION get_municipal_district(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS municipal_districts AS
$$
	SELECT * FROM municipal_districts WHERE id = InItemId;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION get_municipal_districts(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF municipal_districts AS
$$
	SELECT * FROM municipal_districts ORDER BY name;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION create_municipal_district(InNewName VARCHAR, InNewHead VARCHAR, InNewHp VARCHAR, InUserId INTEGER, InPCName VARCHAR, OUT OutNewId INTEGER) AS
$$
BEGIN
	INSERT INTO municipal_districts (name, head, head_position)
	VALUES (InNewName, InNewHead, InNewHp)
	RETURNING id INTO OutNewId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION change_municipal_district(InItemId INTEGER, InNewName VARCHAR, InNewHead VARCHAR, InNewHp VARCHAR, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
BEGIN
	UPDATE municipal_districts
	SET name = InNewName, head = InNewHead, head_position = InNewHp
	WHERE id = InItemId;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION delete_municipal_district(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
DECLARE 
	err_state text;
	err_constraint text;
BEGIN
	DELETE FROM municipal_districts 
	WHERE id = InItemId;
	RETURN;
EXCEPTION WHEN foreign_key_violation THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_villages() RETURNS SETOF villages AS
$$
	SELECT * FROM villages ORDER BY md_id, name;
$$ LANGUAGE SQL STABLE;


CREATE FUNCTION get_village(itemid INTEGER) RETURNS villages AS
$$
	SELECT * FROM villages WHERE id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION change_village(itemid INTEGER, newname VARCHAR(100), newmd INTEGER, newsite VARCHAR(100)) RETURNS VOID AS
$$
BEGIN
	UPDATE villages
	SET name = newname, site_name = newsite, md_id = newmd
	WHERE id = itemid;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION create_village(newname VARCHAR(100), newmd INTEGER, newsite VARCHAR(100), OUT newId INTEGER) AS
$$
BEGIN
	INSERT INTO villages (id, name, md_id, site_name)
	VALUES (DEFAULT, newname, newmd, newsite)
	RETURNING id INTO newId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION delete_village(itemId INTEGER) RETURNS VOID AS
$$
DECLARE 
	err_state text;
	err_constraint text;
BEGIN
	DELETE FROM villages 
	WHERE id = itemId;
	RETURN;
EXCEPTION WHEN foreign_key_violation THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION getStreet(itemid INTEGER) RETURNS streets AS
$$
	SELECT * FROM streets WHERE id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION changeStreet(itemid INTEGER, newname VARCHAR(100), newvillage INTEGER, newsite VARCHAR(100), newtype INTEGER) RETURNS VOID AS
$$
BEGIN
	UPDATE streets
	SET name = newname, site_name = newsite, village_id = newvillage, street_type = newtype
	WHERE id = itemid;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION createStreet(newname VARCHAR(100), newvillage INTEGER, newsite VARCHAR(100), newtype INTEGER, OUT newId INTEGER) AS
$$
BEGIN
	INSERT INTO streets (id, name, village_id, site_name, street_type)
	VALUES (DEFAULT, newname, newvillage, newsite, newtype)
	RETURNING id INTO newId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION deleteStreet(itemId INTEGER) RETURNS VOID AS
$$
DECLARE 
	err_state text;
	err_constraint text;
BEGIN
	DELETE FROM streets
	WHERE id = itemId;
	RETURN;
EXCEPTION WHEN others THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
END;
$$ LANGUAGE plpgsql;


CREATE FUNCTION getDogovor(itemid INTEGER) RETURNS dogovors AS
$$
	SELECT * FROM dogovors WHERE id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION changeDogovor(itemid INTEGER, newname VARCHAR(100), newshortname VARCHAR(30)) RETURNS VOID AS
$$
BEGIN
	UPDATE dogovors
	SET name = newname, short_name = newshortname
	WHERE id = itemid;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION createDogovor(newname VARCHAR(100), newshortname VARCHAR(30), OUT newId INTEGER) AS
$$
BEGIN
	INSERT INTO dogovors (id, name, short_name)
	VALUES (DEFAULT, newname, newshortname)
	RETURNING id INTO newId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION deleteDogovor(itemId INTEGER) RETURNS VOID AS
$$
DECLARE 
	err_state text;
	err_constraint text;
BEGIN
	IF itemId = 0 THEN
	   	RAISE EXCEPTION SQLSTATE '99003';
	ELSE
		DELETE FROM dogovors
		WHERE id = itemId;
	END IF;
	RETURN;
EXCEPTION WHEN others THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_contractors(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF contractors AS
$$
  SELECT * FROM contractors ORDER BY name;
$$ LANGUAGE SQL;

CREATE FUNCTION get_contractor(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS contractors AS
$$
  SELECT * FROM contractors WHERE id = InItemId;
$$ LANGUAGE SQL;

CREATE FUNCTION add_man_hour_rates_to_contractor(InContractorId INTEGER) RETURNS VOID AS
$$
  INSERT INTO man_hour_cost_rates(mode_id, term_id, contractor_id)
  SELECT mhm.id, t.id, InContractorId
  FROM man_hour_cost_modes AS mhm,
  terms AS t
  WHERE t.id = (SELECT id FROM terms WHERE begin_date = (SELECT MAX(begin_date) FROM terms))
  AND mhm.id > 0;
$$ LANGUAGE SQL;
COMMENT ON FUNCTION add_man_hour_rates_to_contractor IS 'Добавление режимов стоимости человекочаса для подрядчика';

CREATE FUNCTION create_contractor(InNewName VARCHAR, InNewDirector VARCHAR, InNewDirPosition VARCHAR, InNewBldnStatus BOOLEAN, InUserId INTEGER, InPCName VARCHAR, OUT newId INTEGER) AS
$$
  BEGIN
    INSERT INTO contractors (id, name, director, bldn_contractor, director_position)
    VALUES (DEFAULT, InNewName, InNewDirector, InNewBldnStatus, InNewDirPosition)
	   RETURNING id INTO newId;

    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 31, InUserId, InPCName, JSONB_AGG(contractors)
      FROM contractors
     WHERE id = newId;

    IF InNewBldnStatus THEN
      PERFORM add_man_hour_rates_to_contractor(newId);
    END IF;
      
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_contractor IS 'Добавление подрядной организации';

CREATE FUNCTION change_contractor(InItemId INTEGER, InNewName VARCHAR, InNewDirector VARCHAR, InNewDirPosition VARCHAR, InNewBldnStatus BOOLEAN, InIsUsing BOOLEAN, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    _bldn_status BOOLEAN;
    _is_using BOOLEAN;

  BEGIN
    IF InItemId = 0 THEN
      RAISE '%,%', get_error_number('delete_denied'), get_error_message('delete_denied');
    END IF;

    IF NOT (InNewBldnStatus AND InIsUsing) THEN
      IF EXISTS (SELECT contractor_id FROM buildings WHERE contractor_id = InItemId) THEN
	RAISE '%,%', get_error_number('has_children'), get_error_message('has_children');
      END IF;
      DELETE FROM man_hour_cost_rates
       WHERE contractor_id = InItemId
	 AND term_id = (SELECT id FROM terms WHERE begin_date = (SELECT MAX(begin_date) FROM terms));
    END IF;

    SELECT bldn_contractor INTO _bldn_status FROM contractors WHERE id = InItemId;
    
    WITH updated_rows AS (
      UPDATE contractors
	 SET name = InNewName,
	     director = InNewDirector,
	     bldn_contractor = InNewBldnStatus,
	     director_position = InNewDirPosition,
	     is_using = InIsUsing
       WHERE id = InItemId
	     RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 32, InUserId, InPCName,
	   JSONB_AGG(ttt)
      FROM (SELECT JSONB_AGG(cont) AS prev, JSONB_AGG(updated_rows) AS upd
	      FROM contractors AS cont, updated_rows
	     WHERE cont.id = InItemId) AS ttt;

    IF InNewBldnStatus AND NOT _bldn_status AND InIsUsing THEN
      PERFORM add_man_hour_rates_to_contractor(InItemId);
    END IF;
    
    RETURN;
END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_contractor IS 'Изменение подрядной организации';

CREATE FUNCTION delete_contractor(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF InItemId = 0 THEN
      RAISE '%,%', get_error_number('delete_denied'), get_error_message('delete_denied');
    END IF;

    WITH deleted_rows AS (
      DELETE FROM contractors
       WHERE id = InItemId
      RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 33, InUserId, InPCName, JSONB_AGG(deleted_rows)
      FROM deleted_rows;
    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_contractor IS 'Удаление подрядной организации';

CREATE FUNCTION getGWT(itemid INTEGER) RETURNS global_work_types AS
$$
	SELECT * FROM global_work_types WHERE id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION changeGWT(itemid INTEGER, newname VARCHAR(50), newdesc TEXT) RETURNS VOID AS
$$
BEGIN
	UPDATE global_work_types
	SET name = newname, description = newdesc
	WHERE id = itemid;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION createGWT(newname VARCHAR(50), newdesc TEXT, OUT newId INTEGER) AS
$$
BEGIN
	INSERT INTO global_work_types (id, name, description)
	VALUES (DEFAULT, newname, newdesc)
	RETURNING id INTO newId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION deleteGWT(itemId INTEGER) RETURNS VOID AS
$$
DECLARE 
	err_state text;
	err_constraint text;
BEGIN
	DELETE FROM global_work_types
	WHERE id = itemId;
	RETURN;
EXCEPTION WHEN others THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
END;
$$ LANGUAGE plpgsql;


CREATE FUNCTION getImprovement(itemid INTEGER) RETURNS improvements AS
$$
	SELECT * FROM improvements WHERE id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION changeImprovement(itemid INTEGER, newname VARCHAR(100), newshortname VARCHAR(30)) RETURNS VOID AS
$$
BEGIN
	IF itemid = 0 THEN
	   	RAISE EXCEPTION SQLSTATE '99003';
	ELSE
		UPDATE improvements
		SET name = newname, short_name = newshortname
		WHERE id = itemid;
	END IF;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION createImprovement(newname VARCHAR(100), newshortname VARCHAR(30), OUT newId INTEGER) AS
$$
BEGIN
	INSERT INTO improvements (id, name, short_name)
	VALUES (DEFAULT, newname, newshortname)
	RETURNING id INTO newId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION deleteImprovement(itemId INTEGER) RETURNS VOID AS
$$
DECLARE 
	err_state text;
	err_constraint text;
BEGIN
	IF itemId = 0 THEN
	   	RAISE EXCEPTION SQLSTATE '99003';
	ELSE
		DELETE FROM improvements
		WHERE id = itemId;
	END IF;
	RETURN;
EXCEPTION WHEN others THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
END;
$$ LANGUAGE plpgsql;


CREATE FUNCTION getPlanStatus(itemid INTEGER) RETURNS plan_work_statuses AS
$$
	SELECT * FROM plan_work_statuses WHERE id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION getWallMaterial(itemid INTEGER) RETURNS wall_materials AS
$$
	SELECT * FROM wall_materials WHERE id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION changeWallMaterial(itemid INTEGER, newname VARCHAR(100)) RETURNS VOID AS
$$
BEGIN
	IF itemid = 0 THEN
	   	RAISE EXCEPTION SQLSTATE '99003';
	ELSE
		UPDATE wall_materials
		SET name = newname
		WHERE id = itemid;
	END IF;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION createWallMaterial(newname VARCHAR(100), OUT newId INTEGER) AS
$$
BEGIN
	INSERT INTO wall_materials (id, name)
	VALUES (DEFAULT, newname)
	RETURNING id INTO newId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION deleteWallMaterial(itemId INTEGER) RETURNS VOID AS
$$
DECLARE 
	err_state text;
	err_constraint text;
BEGIN
	IF itemId = 0 THEN
	   	RAISE EXCEPTION SQLSTATE '99003';
	ELSE
		DELETE FROM wall_materials
		WHERE id = itemId;
	END IF;
	RETURN;
EXCEPTION WHEN others THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
END;
$$ LANGUAGE plpgsql;


CREATE FUNCTION getWorkType(itemid INTEGER) RETURNS work_types AS
$$
	SELECT * FROM work_types WHERE id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION changeWorkType(itemid INTEGER, newname VARCHAR(100)) RETURNS VOID AS
$$
BEGIN
	UPDATE work_types
	SET name = newname
	WHERE id = itemid;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION createWorkType(newname VARCHAR(100), OUT newId INTEGER) AS
$$
BEGIN
	INSERT INTO work_types (id, name)
	VALUES (DEFAULT, newname)
	RETURNING id INTO newId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION deleteWorkType(itemId INTEGER) RETURNS VOID AS
$$
DECLARE 
	err_state text;
	err_constraint text;
BEGIN
	DELETE FROM work_types
	WHERE id = itemId;
	RETURN;
EXCEPTION WHEN others THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
END;
$$ LANGUAGE plpgsql;


CREATE FUNCTION getWorkKind(itemid INTEGER) RETURNS work_kinds AS
$$
	SELECT * FROM work_kinds WHERE id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION changeWorkKind(itemid INTEGER, newname VARCHAR(300), newwt INTEGER) RETURNS VOID AS
$$
DECLARE
	iid INTEGER;
BEGIN
	SELECT MIN(id) INTO iid FROM work_kinds WHERE lower(name) = lower(newname) AND worktype_id = newwt AND id <> itemid;
	IF iid IS NOT NULL THEN
	   RAISE EXCEPTION 'work_kind_in_work_type_must_be_unique';
	ELSE
		UPDATE work_kinds
		SET name = newname, worktype_id = newwt
		WHERE id = itemid;
	END IF;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION createWorkKind(newname VARCHAR(300), newwt INTEGER, OUT newId INTEGER) AS
$$
DECLARE
	iid INTEGER;
BEGIN
	SELECT MIN(id) INTO iid FROM work_kinds WHERE lower(name) = lower(newname) AND worktype_id = newwt;
	IF iid IS NOT NULL THEN
	   RAISE EXCEPTION 'work_kind_in_work_type_must_be_unique';
	ELSE
		INSERT INTO work_kinds (id, name, worktype_id)
		VALUES (DEFAULT, newname, newwt)
		RETURNING id INTO newId;
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION deleteWorkKind(itemId INTEGER) RETURNS VOID AS
$$
DECLARE 
	err_state text;
	err_constraint text;
BEGIN
	DELETE FROM work_kinds
	WHERE id = itemId;
	RETURN;
EXCEPTION WHEN others THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION getWorkKindsByWT(wtId INTEGER DEFAULT -1001) RETURNS SETOF work_kinds AS
$$
BEGIN
	IF wtId = -1001 THEN
	   RETURN QUERY
	   SELECT * FROM work_kinds
	   ORDER BY worktype_id, name;
	ELSE
	   RETURN QUERY
	   SELECT * FROM  work_kinds
	   WHERE worktype_id = wtId
	   ORDER BY name;
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_mc(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS management_companies AS
$$
	SELECT * FROM management_companies WHERE id = InItemId;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION changeMC(itemid INTEGER, newname VARCHAR(50), newreport VARCHAR(15), newnotmanage BOOLEAN) RETURNS VOID AS
$$
BEGIN
	IF itemid = 0 THEN
	   	RAISE EXCEPTION SQLSTATE '99003';
	ELSE
		UPDATE management_companies
		SET name = newname, report_name = newreport, not_manage = newnotmanage
		WHERE id = itemid;
	END IF;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION createMC(newname VARCHAR(50), newreport VARCHAR(15), newnotmanage BOOLEAN, OUT newId INTEGER) AS
$$
BEGIN
	INSERT INTO management_companies (id, name, report_name, not_manage)
	       VALUES (DEFAULT, newname, newreport, newnotmanage)
	       RETURNING id INTO newId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION deleteMC(itemId INTEGER) RETURNS VOID AS
$$
DECLARE 
	err_state text;
	err_constraint text;
BEGIN
	IF itemId = 0 THEN
	   RAISE EXCEPTION SQLSTATE '99003';
	ELSE
		DELETE FROM management_companies
		WHERE id = itemId;
	END IF;
	RETURN;
EXCEPTION WHEN others THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION getEmployee(itemid INTEGER) RETURNS employees AS
$$
	SELECT * FROM employees WHERE id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION changeEmployee(itemid INTEGER, newfname VARCHAR(60), newsname VARCHAR(60), newlname VARCHAR(60), neworgid INTEGER, newpos INTEGER, newposname VARCHAR(100), newsign BOOLEAN) RETURNS VOID AS
$$
DECLARE
	empid INTEGER;
BEGIN
	IF newpos = 1 OR newpos = 2 THEN
		SELECT MIN(id) INTO empid FROM employees WHERE position_status = newpos AND organization_id = neworgid AND id <> itemid;
	END IF;
	IF empid IS NOT NULL THEN	
		RAISE EXCEPTION 'director_or_chief_engeneer_must_be_unique';
	ELSE
		UPDATE employees
		SET organization_id = neworgid,
	    	    last_name = newlname,
	    	    first_name = newfname,
	    	    second_name = newsname,
	    	    position_status = newpos,
	    	    position_name = newposname,
	    	    sign_report = newsign
		WHERE id = itemid;
	END IF;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION createEmployee(newfname VARCHAR(60), newsname VARCHAR(60), newlname VARCHAR(60), neworgid INTEGER, newpos INTEGER, newposname VARCHAR(100), newsign BOOLEAN, OUT newid INTEGER) AS
$$
DECLARE
	empid INTEGER;
BEGIN
	IF newpos = 1 OR newpos = 2 THEN
		SELECT MIN(id) INTO empid FROM employees WHERE position_status = newpos AND organization_id = neworgid;
	END IF;
	IF empid IS NOT NULL THEN
		RAISE EXCEPTION 'director_or_chief_engeneer_must_be_unique';
	ELSE
		INSERT INTO employees (id, organization_id, last_name, first_name, second_name, position_status, position_name, sign_report)
	       	VALUES (DEFAULT, neworgid, newlname, newfname, newsname, newpos, newposname, newsign)
	       	RETURNING id INTO newId;
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION deleteEmployee(itemId INTEGER) RETURNS VOID AS
$$
DECLARE 
	err_state text;
	err_constraint text;
BEGIN
	DELETE FROM employees
	WHERE id = itemId;
	RETURN;
EXCEPTION WHEN others THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION getEmployeesInOrganization(orgId INTEGER) RETURNS SETOF employees AS
$$
	SELECT * FROM employees WHERE organization_id = orgId;
$$ LANGUAGE SQL;

CREATE FUNCTION get_plan_statuses(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF plan_work_statuses AS
$$
  SELECT * FROM plan_work_statuses ORDER BY name;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION get_plan_statuses_new_work(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF plan_work_statuses AS
$$
  SELECT * FROM plan_work_statuses WHERE can_new;
$$ LANGUAGE SQL STABLE;


CREATE FUNCTION get_plan_work(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS plan_works AS
$$
  DECLARE
    _out_row plan_works%ROWTYPE;
  BEGIN
    SELECT * INTO _out_row FROM plan_works WHERE id = InItemId;

    IF NOT user_has_right_read(InUserId, rights_get_plan_rights_number(_out_row.gwt_id)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
  
    RETURN _out_row;
  END;
$$ LANGUAGE plpgsql STABLE;

CREATE FUNCTION create_plan_work(InBldn INTEGER, InMc INTEGER, InGwt INTEGER, InWk INTEGER, InDate DATE, InSum NUMERIC(10,2), InSmetaSum NUMERIC(10, 2), InNote TEXT, InPrivateNote TEXT, InContractor INTEGER, InStatus INTEGER, InEmployee VARCHAR(200), InUserId INTEGER, InPcName VARCHAR(100), OUT OutId INTEGER) AS
$$
BEGIN
  IF NOT user_has_right_change(InUserId, rights_get_plan_rights_number(InGwt)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
  END IF;

  INSERT INTO plan_works(id, gwt_id, workkind_id, bldn_id, work_date, work_sum, note, private_note, contractor_id, mc_id, work_status, employee, smeta_sum, create_user)
  VALUES (DEFAULT, InGwt, InWk, InBldn, InDate, InSum, InNote, InPrivateNote, InContractor, InMc, InStatus, InEmployee, InSmetaSum, InUserId)
	 RETURNING id INTO OutId;
  PERFORM add_log_action(4, InUserId, InPcName, log_plan_work_string(OutId));
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION change_plan_work(InItemId INTEGER, InGwt INTEGER, InWk INTEGER, InDate DATE, InSum NUMERIC(10,2), InSmetaSum NUMERIC(10,2), InNote TEXT, InPrivateNote TEXT, InContractor INTEGER, InStatus INTEGER, InEmployee VARCHAR(200), InWR INTEGER, InBDate DATE, InEDate DATE, InUserId INTEGER, InPcName VARCHAR) RETURNS VOID AS
$$
DECLARE
	work_string TEXT;
BEGIN
  IF NOT user_has_right_change(InUserId, rights_get_plan_rights_number(InGwt)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
  END IF;

  SELECT 'Было: ' || log_plan_work_string(InItemId) INTO work_string;
  UPDATE plan_works
     SET gwt_id = InGwt,
	 workkind_id = InWk,
	 work_date = InDate,
	 work_sum = InSum,
	 note = InNote,
	 private_note = InPrivateNote,
	 contractor_id = InContractor,
	 work_status = InStatus,
	 employee = InEmployee,
	 work_ref = InWR,
	 begin_date = InBDate,
	 end_date = InEDate,
	 smeta_sum = InSmetaSum,
	 last_change_user = InUserId
   WHERE id = InItemId;
	SELECT work_string || ' Стало: ' || log_plan_work_string(InItemId) INTO work_string;
	PERFORM add_log_action(5, InUserId, InPcName, work_string);
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION delete_plan_work(InItemId INTEGER, InUserId INTEGER, InPcName VARCHAR(100)) RETURNS VOID AS
$$
  DECLARE 
    err_state TEXT;
    err_constraint TEXT;
    work_string TEXT;
    workgwt INTEGER;
  BEGIN
    SELECT gwt_id INTO workgwt FROM plan_works WHERE id = InItemId;
    
    IF NOT user_has_right_delete(InUserId, rights_get_plan_rights_number(workgwt)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
    
    SELECT log_plan_work_string(InItemId) INTO work_string;
    DELETE FROM plan_works
     WHERE id = InItemId;
    PERFORM add_log_action(6, InUserId, InPcName, work_string);
    RETURN;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_bldn_plan_years(inBldnId integer) RETURNS SETOF integer AS
$$
BEGIN
	RETURN QUERY
	SELECT DISTINCT (EXTRACT ( YEAR FROM work_date ))::integer
	FROM plan_works
	WHERE bldn_id = inBldnId
	ORDER BY 1 DESC;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_plan_works_by_bldn(inBldnId INTEGER, inBeginDate date, inEndDate date, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF plan_works AS
$$
BEGIN
  RETURN QUERY
    SELECT pw.* FROM plan_works AS pw
    INNER JOIN plan_work_statuses AS pws ON pw.work_status = pws.id
    WHERE bldn_id = inBldnId 
    AND (work_date >= inBeginDate OR inBeginDate IS NULL)
    AND (work_date <= inEndDate OR inEndDate IS NULL)
    ORDER BY work_date DESC;
END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_plan_works_by_bldn IS 'Планируемые работы в доме';

CREATE FUNCTION open_next_term() RETURNS VOID AS
$$
  DECLARE
    _last_term_date DATE;
    _new_begin_date DATE;
    _new_end_date DATE;
    _last_term_id INTEGER;
  BEGIN
    SELECT MAX(id) INTO _last_term_id FROM terms;
    
    IF _last_term_id IS NULL THEN
      _new_begin_date := make_date(EXTRACT(year FROM current_date)::INTEGER, EXTRACT(MONTH FROM current_date)::INTEGER, 1);
      _last_term_id := 0;
    ELSE
      SELECT MAX(begin_date) INTO _last_term_date FROM terms;
      _new_begin_date := _last_term_date + INTERVAL '1 month';
    END IF;
    _new_end_date := _new_begin_date + INTERVAL '1 month' - INTERVAL '1 day';
    
    INSERT INTO terms (id, begin_date, end_date)
    VALUES(_last_term_id + 1, _new_begin_date, _new_end_date);
    
    IF _last_term_id > 0 THEN 
      -- копирование стоимости человекочаса
      INSERT INTO man_hour_cost_rates(mode_id, term_id, contractor_id, cost_sum)
      SELECT mode_id, _last_term_id+1, contractor_id, cost_sum
      FROM man_hour_cost_rates
      WHERE term_id = _last_term_id;

      -- копирование режима стоимости человекочаса в домах
      INSERT INTO bldn_man_hour_cost(term_id, bldn_id, mode_id)
      SELECT _last_term_id+1, bldn_id, mode_id
      FROM bldn_man_hour_cost
      WHERE term_id = _last_term_id;

      -- копирование режимов услуг в доме
      INSERT INTO bldn_services_history
      SELECT _last_term_id, * FROM bldn_services;
      
      -- архивирование информации о домах
      INSERT INTO buildings_history
      SELECT _last_term_id, * FROM buildings;
    
      INSERT INTO buildings_tech_info_history
      SELECT _last_term_id, * FROM buildings_tech_info;
    
      INSERT INTO buildings_land_info_history
      SELECT _last_term_id, * FROM buildings_land_info;

      -- история информации об элементах общего имущества и их параметрах
      INSERT INTO building_common_property_elements_history
      SELECT _last_term_id, * FROM building_common_property_elements;

      INSERT INTO building_common_property_element_parameter_history
      SELECT _last_term_id, * FROM building_common_property_element_parameter;

    END IF;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION open_next_term IS 'Открытие нового месяца';

CREATE FUNCTION get_bldn_types(userId INTEGER) RETURNS SETOF bldn_types AS
$$
	SELECT * FROM bldn_types ORDER BY id;
$$ LANGUAGE SQL;

CREATE FUNCTION get_building(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS buildings AS
$$
  SELECT * FROM buildings WHERE id = InItemId;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_building IS 'Информация о доме по коду';

CREATE FUNCTION create_building(newbldn INTEGER, newstreet INTEGER, newbldnno VARCHAR(10), newmc INTEGER, newcontract INTEGER, InUserId INTEGER, InPCName VARCHAR, OUT newid INTEGER) AS
$$
  BEGIN
    IF newbldn = 0 THEN
      INSERT INTO buildings(id, street_id, bldn_no, mc_id, dogovor_type) VALUES (DEFAULT, newstreet, newbldnno, newmc, newcontract) RETURNING id INTO newid;
    ELSE
      INSERT INTO buildings(id, street_id, bldn_no, mc_id, dogovor_type) VALUES (newbldn, newstreet, newbldnno, newmc, newcontract) RETURNING id INTO newid;
    END IF;

    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 100, InUserId, InPCName, JSONB_AGG(buildings)
      FROM buildings
     WHERE id = newid;

    RETURN;
END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_building IS 'Создание дома';

CREATE FUNCTION change_bldn_services(InItemId INTEGER, InHw INTEGER, InCw INTEGER, InGas INTEGER, InHeating INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    WITH ut AS (
      UPDATE buildings
	 SET hot_water = InHw,
	     cold_water = InCw,
	     heating = InHeating,
	     gas = InGas
       WHERE id = InItemId
	     RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 102, InUserId, InPCName,
	   JSONB_AGG(ttt)
      FROM (SELECT JSONB_AGG(ot) AS prev, JSONB_AGG(uut) AS upd
	      FROM (SELECT id, hot_water, cold_water, heating, gas FROM buildings WHERE id = InItemId) AS ot,
		   (SELECT id, hot_water, cold_water, heating, gas FROM ut) AS uut)
	     AS ttt;
    
    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_bldn_services IS 'Изменение услуг в доме';

CREATE FUNCTION change_bldn_common(itemid INTEGER, newimp INTEGER, newtype INTEGER, newsite VARCHAR(10), newcadastral VARCHAR(20), newdisrepair BOOLEAN, newenergoclass INTEGER, newfias VARCHAR, newgisguid VARCHAR, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    WITH ut AS (
      UPDATE buildings
	 SET improvement_id = newimp,
	     bldn_type = newtype,
	     site_no = newsite,
	     cadastral_no = newcadastral,
	     disrepair = newdisrepair,
	     energo_class = newenergoclass,
	     fias = newfias,
	     gis_guid = newgisguid
       WHERE id = itemid
	     RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
	SELECT 103, InUserId, InPCName, JSONB_AGG(ttt)
	  FROM (SELECT JSONB_AGG(ot) AS prev, JSONB_AGG(uut) AS upd
		  FROM	    
		    (SELECT id, improvement_id, bldn_type, site_no, cadastral_no, disrepair, energo_class, fias, gis_guid FROM buildings WHERE id = itemid) AS ot,
		    (SELECT id, improvement_id, bldn_type, site_no, cadastral_no, disrepair, energo_class, fias, gis_guid FROM ut) AS uut)
		 AS ttt;

	RETURN;
END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_bldn_common IS 'Изменение общей информации дома'; 

CREATE FUNCTION change_bldn_dogovor(itemId INTEGER, newmc INTEGER, newcontractor INTEGER, newdogovor INTEGER, newout BOOLEAN, NewManHourMode INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    _term_id INTEGER;
    _old_contractor INTEGER;
  BEGIN

    SELECT id INTO _term_id FROM terms WHERE begin_date = (SELECT MAX(begin_date) FROM terms);
    SELECT contractor_id INTO _old_contractor FROM buildings WHERE id = itemId;

    WITH ut AS (
      UPDATE buildings
	 SET mc_id = newmc,
	     contractor_id = newcontractor,
	     dogovor_type = newdogovor,
	     out_report = newout
       WHERE id = itemId
	     RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
	SELECT 103, InUserId, InPCName, JSONB_AGG(ttt)
	  FROM (SELECT JSONB_AGG(ot) AS prev, JSONB_AGG(uut) AS upd
		  FROM	    
		    (SELECT id, mc_id, contractor_id, dogovor_type, out_report FROM buildings WHERE id = itemId) AS ot,
		    (SELECT id, mc_id, contractor_id, dogovor_type, out_report FROM ut) AS uut)
		 AS ttt;

    IF newcontractor = 0 THEN 
      -- Если новый подрядчик == 0, то удаляем режим человекочаса

      IF _old_contractor != 0 THEN
	IF EXISTS (SELECT * FROM works WHERE gwt_id = 1 AND bldn_id = itemId AND work_date = _term_id) THEN
	  RAISE '%, %', get_error_number('has_children'), get_error_message('has_children');
	END IF;
	WITH deleted_rows AS (
	  DELETE FROM bldn_man_hour_cost
	   WHERE bldn_id = itemId
	     AND term_id = _term_id
	  RETURNING *
	)
	    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
	SELECT 30, InUserId, InPCName, JSONB_AGG(deleted_rows) AS prev
	  FROM deleted_rows;
      END IF;			-- _old_contractor == 0

    ELSE 			-- newcontractor = 0
		  
      IF _old_contractor = 0 THEN
	-- Если не было подрядчика, то добавляем режим человекочаса
	WITH inserted_rows AS (
	  INSERT INTO bldn_man_hour_cost(term_id, bldn_id, mode_id)
	  VALUES (_term_id, itemId, 0)
		 RETURNING *
	)
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
	SELECT 30, InUserId, InPCName, JSONB_AGG(inserted_rows) AS upd
	FROM inserted_rows;
      ELSE 			-- _old_contractor == 0
	WITH ut AS (
	  UPDATE bldn_man_hour_cost
	     SET mode_id = NewManHourMode
	   WHERE bldn_id = itemId
	     AND term_id = _term_id
		 RETURNING *
	)
	    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
	SELECT 30, InUserId, InPCName, JSONB_AGG(ttt)
	  FROM (SELECT JSONB_AGG(ot) AS prev, JSONB_AGG(ut) AS upd
		  FROM	    
		    (SELECT * FROM bldn_man_hour_cost WHERE bldn_id = itemId AND term_id = _term_id) AS ot,
		    ut)
		 AS ttt;
      END IF;			-- _old_contractor == 0
    END IF;			-- newcontractor == 0

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_bldn_dogovor IS 'Изменение информации о договорах дома';

CREATE FUNCTION delete_building(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    WITH deleted_rows AS (
      DELETE FROM buildings
       WHERE id = InItemId
      RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 101, InUserId, InPCName, JSONB_AGG(deleted_rows)
      FROM deleted_rows;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_building IS 'Удаление дома';

CREATE FUNCTION bldn_address(bldnId INTEGER, OUT address TEXT) AS
$$
BEGIN
	SELECT xs.name || ' д.' || b.bldn_no INTO address
	FROM buildings b
	     INNER JOIN xstreets xs ON b.street_id = xs.id
	WHERE b.id = bldnId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_bldn_tech_info(itemid INTEGER) RETURNS buildings_tech_info AS
$$
	SELECT * FROM buildings_tech_info WHERE bldn_id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION change_bldn_tech_info(itemid INTEGER, newfmin INTEGER, newfmax INTEGER, newvaults INTEGER, newentrance INTEGER, newstairs INTEGER, newbuilt INTEGER, newcommissioning INTEGER, newdepreciation REAL, newatticsq REAL, newvaultssq REAL, newstairssq REAL, newcorridorsq REAL, newothersq REAL, newstrvolume REAL, newwall INTEGER, newhaselectro BOOLEAN, newhashw BOOLEAN, newhascommon BOOLEAN, newhascw BOOLEAN, newhasheating BOOLEAN, newhasdoorphone BOOLEAN, newdoorphonecomment TEXT, newhasthermoregulator BOOLEAN, InBanisterSq REAL, InDoorsSq REAL, InWindowSillsSq REAL, InDoorHandlesSq REAL, InMailBoxesSq REAL, InRadiatorsSq REAL, InHasDoorCloser BOOLEAN, InUserId INTEGER, InPcName VARCHAR) RETURNS VOID AS
$$
BEGIN
  IF NOT user_has_right_change(InUserId, rights_get_tech_info_rights_number()) THEN
    RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
  END IF;

  WITH tt AS (
    UPDATE buildings_tech_info
       SET floor_min = newfmin,
	   floor_max = newfmax,
	   vaults = newvaults,
	   entrances = newentrance,
	   stairs = newstairs,
	   built_year = newbuilt,
	   commissioning_year = newcommissioning,
	   depreciation = newdepreciation,
	   attic_square = newatticsq,
	   vaults_square = newvaultssq,
	   stairs_square = newstairssq,
	   corridor_square = newcorridorsq,
	   other_square = newothersq,
	   structural_volume = newstrvolume,
	   wallmater_id = newwall,
	   has_odpu_electro = newhaselectro,
	   has_odpu_hotwater = newhashw,
	   has_odpu_common = newhascommon,
	   has_odpu_heating = newhasheating,
	   has_odpu_coldwater = newhascw,
	   has_doorphone = newhasdoorphone,
	   doorphone_comment = newdoorphonecomment,
	   has_thermoregulator = newhasthermoregulator,
	   square_banisters = InBanisterSq,
	   square_doors = InDoorsSq,
	   square_windowsills = InWindowSillsSq,
	   square_doorhandles = InDoorHandlesSq,
	   square_mailboxes = InMailBoxesSq,
	   square_radiators = InRadiatorsSq,
	   has_doorcloser = InHasDoorCloser
     WHERE bldn_id = itemid
	   RETURNING *
  )
      INSERT INTO log_log(action_id, user_id, pc_name, log_action)
  SELECT 105, InUserId, InPCName,
	 JSONB_AGG(ttt)
    FROM (SELECT JSONB_AGG(bti) AS prev, JSONB_AGG(tt) AS upd
	    FROM buildings_tech_info AS bti, tt
	   WHERE bti.bldn_id = itemId) AS ttt;
  
  RETURN;
END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_bldn_tech_info IS 'Изменение технической информации МКД';

CREATE FUNCTION getBuildingLandInfo(itemid INTEGER) RETURNS buildings_land_info AS
$$
	SELECT * FROM buildings_land_info WHERE bldn_id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION change_bldn_land_info(itemid INTEGER, newinv REAL, newuse REAL, newsurv REAL, newbuilt REAL, newundev REAL, newhard REAL, newdrive REAL, newside REAL, newother REAL, newcadastr VARCHAR(20), newsaf BOOLEAN, newfences BOOLEAN) RETURNS VOID AS
$$
BEGIN
	UPDATE buildings_land_info
	SET inventory_area = newinv,
	    use_area = newuse,
	    survey_area = newsurv,
	    builtup_area = newbuilt,
	    undeveloped_area = newundev,
	    hard_coatings = newhard,
	    drive_ways_hard = newdrive,
	    side_walks_hard = newside,
	    others_hard = newother,
	    cadastral_no = newcadastr,
	    saf = newsaf,
	    fences = newfences
	WHERE bldn_id = itemid;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_fsource(InItemId INTEGER) RETURNS work_financing_sources AS
$$
	SELECT * FROM work_financing_sources WHERE id = InItemId;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION get_fsources() RETURNS SETOF work_financing_sources AS
$$
	SELECT * FROM work_financing_sources ORDER BY name;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION change_fsource(InItemId INTEGER, InName VARCHAR(50), InFromSubaccount BOOLEAN, InNote TEXT) RETURNS VOID AS
$$
BEGIN
	IF InItemId = 1 THEN
	   	RAISE EXCEPTION SQLSTATE '99003';
	ELSE
		UPDATE work_financing_sources
		SET name = InName, note = InNote, from_subaccount = InFromSubaccount
		WHERE id = InItemId;
	END IF;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION create_fsource(InName VARCHAR(50), InNote TEXT, InFromSubaccount BOOLEAN, OUT OutId INTEGER) AS
$$
BEGIN
	INSERT INTO work_financing_sources (name, from_subaccount, note)
	VALUES (InName, InFromSubaccount, InNote)
	RETURNING id INTO OutId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION delete_fsource(InItemId INTEGER) RETURNS VOID AS
$$
DECLARE 
	err_state text;
	err_constraint text;
BEGIN
	IF InItemId <= 1 THEN
	   	RAISE EXCEPTION SQLSTATE '99003';
	ELSE
		DELETE FROM work_financing_sources
		WHERE id = InItemId;
	END IF;
	RETURN;
EXCEPTION WHEN foreign_key_violation THEN
	RAISE EXCEPTION '%, нельзя удалять таблицу, т.к. на неё есть ссылки', SQLSTATE;
END;
$$ LANGUAGE plpgsql;


CREATE FUNCTION log_work_string(wId INTEGER, OUT wstr TEXT) AS
$$
  BEGIN
    SELECT CONCAT_WS(' :-: ', w.id, w.bldn_id, bldn_address(w.bldn_id), mc.name, c.name, gwt.name, wk.name, w.work_sum, t.begin_date, w.volume, w.si, w.note, w.private_note, w.dogovor, fs.name, w.print_flag) INTO wstr
      FROM works w
	   INNER JOIN management_companies mc ON w.mc_id = mc.id
	   INNER JOIN contractors c ON w.contractor_id = c.id
	   INNER JOIN global_work_types gwt ON w.gwt_id = gwt.id
	   INNER JOIN work_kinds wk ON w.workkind_id = wk.id
	   INNER JOIN terms t ON t.id = w.work_date
	   INNER JOIN work_financing_sources fs ON w.finance_source = fs.id
     WHERE w.id = wId;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION log_plan_work_string(wId INTEGER, OUT wstr TEXT) AS
$$
  BEGIN
    SELECT CONCAT_WS(' :-: ', w.id, w.bldn_id, bldn_address(w.bldn_id), mc.name,
		     c.name, gwt.name, wk.name, w.work_sum, w.smeta_sum,
		     w.work_date, w.note, w.private_note, pws.name, w.employee, w.begin_date, w.end_date) INTO wstr
      FROM plan_works w
	   INNER JOIN management_companies mc ON w.mc_id = mc.id
	   INNER JOIN contractors c ON w.contractor_id = c.id
	   INNER JOIN global_work_types gwt ON w.gwt_id = gwt.id
	   INNER JOIN work_kinds wk ON w.workkind_id = wk.id
	   INNER JOIN plan_work_statuses pws ON w.work_status = pws.id
     WHERE w.id = wId;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION add_log_action(actId INTEGER, uId INTEGER, pcName VARCHAR(100), actDesc TEXT) RETURNS VOID AS
$$
BEGIN
	INSERT INTO log_log(action_id, user_id, pc_name, action_description)
	VALUES (actId, uId, pcName, actDesc);
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_work(itemid INTEGER) RETURNS works AS
$$
	SELECT * FROM works WHERE id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION create_work(newgwt INTEGER, newwk INTEGER, newdate INTEGER, newsum NUMERIC(10,2), newsi VARCHAR(20), newvolume VARCHAR(50), newnote TEXT, newpnote TEXT, newcontractor INTEGER, newmc INTEGER, newdogovor VARCHAR(200), newfsource INTEGER, newpf BOOLEAN, newbldn INTEGER, userId INTEGER, pcName VARCHAR, OUT newId INTEGER) AS
$$
  BEGIN
    IF NOT user_has_right_change(userId, rights_get_work_rights_number(newgwt)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    INSERT INTO works(id, gwt_id, workkind_id, bldn_id, work_date, work_sum, si, volume, note, contractor_id, mc_id, dogovor, finance_source, print_flag, private_note)
    VALUES (DEFAULT, newgwt, newwk, newbldn, newdate, newsum, newsi, newvolume, newnote, newcontractor, newmc, newdogovor, newfsource, newpf, newpnote)
	   RETURNING id INTO newId;

    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 1, userId, pcName, JSON_AGG(works)
      FROM works
     WHERE id = newId;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_work IS 'Добавление работы';

CREATE FUNCTION change_work(itemid INTEGER, newgwt INTEGER, newwk INTEGER, newdate INTEGER, newsum NUMERIC(10,2), newsi VARCHAR(20), newvolume VARCHAR(50), newnote TEXT, newpnote TEXT, newcontractor INTEGER, newmc INTEGER, newdogovor VARCHAR(200), newfsource INTEGER, newpf BOOLEAN, userId INTEGER, pcName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_change(userId, rights_get_work_rights_number(newgwt)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
    
    WITH up AS (
      UPDATE works
	 SET gwt_id = newgwt,
	     workkind_id = newwk,
	     work_date = newdate,
	     work_sum = newsum,
	     si = newsi,
	     volume = newvolume,
	     note = newnote,
	     private_note = newpnote,
	     contractor_id = newcontractor,
	     mc_id = newmc,
	     dogovor = newdogovor,
	     finance_source = newfsource,
	     print_flag = newpf,
	     change_date = CURRENT_TIMESTAMP
       WHERE id = itemid
	     RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 2, userId, pcName, JSONB_AGG(ttt)
      FROM (SELECT JSONB_AGG(pr) AS prev, JSONB_AGG(up) AS upd
	      FROM works AS pr, up
	     WHERE pr.id = itemid) AS ttt;
    RETURN;
END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_work IS 'Изменение работы';

CREATE FUNCTION delete_work(itemId INTEGER, userId INTEGER, pcName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    _planwork INTEGER;
    _workgwt INTEGER;
    _maintenanceid INTEGER;
  BEGIN
    SELECT gwt_id INTO _workgwt FROM works WHERE id = itemId;
    IF NOT user_has_right_delete(userId, rights_get_work_rights_number(_workgwt)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    SELECT id INTO _planwork FROM plan_works WHERE work_ref = itemId;
    IF _planwork IS NOT NULL THEN RAISE EXCEPTION SQLSTATE '99003' USING MESSAGE = 'Нельзя удалить запись, т.к. на нее есть ссылки (плановая работа)'; END IF;

    SELECT id INTO _maintenanceid FROM hidden_maintenance_works WHERE workref_id = itemId;
    IF _maintenanceid IS NOT NULL THEN
      PERFORM delete_maintenance_work(_maintenanceid, userId, pcName);
    END IF;

    WITH deleted_row AS (
      DELETE FROM works
       WHERE id = itemId
      RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 3, userId, pcName, JSONB_AGG(deleted_row)
      FROM deleted_row;

    RETURN;
END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_work IS 'Удаление работы';

CREATE VIEW certificate_tmp_counters AS
  WITH tmp_certificates AS (
    SELECT *,
	   RANK() OVER (PARTITION BY counter_id ORDER BY CAST((certificate_date + MAKE_INTERVAL(years => certificate_validite)) AS DATE) DESC)
      FROM certificates),
  max_certificates AS (
    SELECT * from tmp_certificates WHERE rank=1)
  SELECT tc.id AS id,
	 tc.bldn_id AS bldn_id,
	 bldn_address(tc.bldn_id) AS address,
	 tc.name AS name,
	 c.certificate_date,
	 c.certificate_date + MAKE_INTERVAL(years => c.certificate_validite) as end_date
    FROM tmp_counters tc
	   LEFT JOIN max_certificates c ON c.counter_id = tc.id;

CREATE VIEW managed_buildings AS
       SELECT * FROM buildings
       WHERE dogovor_type > 0;

CREATE VIEW bldn_id_no_list AS
  SELECT
    concat_ws(' ', xs.name, b.bldn_no) AS address,
    b.id AS bid,
    b.bldn_no,
    b.out_report,
    b.mc_id, 
    CASE WHEN ASCII(RIGHT(b.bldn_no, 1)) > 57 THEN LPAD(b.bldn_no, 10, '0')
      WHEN STRPOS(b.bldn_no, '.') > 0 THEN LPAD(CONCAT(SUBSTR(b.bldn_no, 1, STRPOS(b.bldn_no, '.')-1), RIGHT(b.bldn_no, 1)), 10, '0')
      ELSE LPAD(b.bldn_no, 9, '0') || '0' END AS tmpAdr,
    xs.site_name || b.site_no AS site_name,
    xs.id AS street_id,
    xs.vid AS village_id,
    xs.mid AS md_id,
    b.dogovor_type
    FROM buildings b
	 INNER JOIN xstreets AS xs ON b.street_id = xs.id
   ORDER BY md_id, xs.name, tmpAdr;
COMMENT ON VIEW bldn_id_no_list IS 'Список домов с адресом';

CREATE VIEW buildings_workslist AS
       SELECT w.id AS id,
       wt.name AS worktype_name,
       wk.name AS workkind_name,
       c.name AS contractor_name,
       w.dogovor AS dogovor,
       t.begin_date AS workdate,
       w.work_sum AS worksum,
       w.volume AS volume,
       w.note AS note,
       w.print_flag AS print_flag,
       w.bldn_id AS bldn_id,
       w.gwt_id,
       w.si AS si,
       fs.name AS fsource,
       t.id AS term_id
       FROM works w
       	    INNER JOIN work_kinds wk ON w.workkind_id = wk.id
	    INNER JOIN work_types wt ON wk.worktype_id = wt.id
	    INNER JOIN contractors c ON w.contractor_id = c.id
	    INNER JOIN work_financing_sources fs ON w.finance_source = fs.id
	    INNER JOIN terms t ON t.id = w.work_date;

CREATE VIEW maintenance_works AS
       SELECT mw.id,
	      mw.man_hours,
	      mw.man_hour_mode_id,
	      w.contractor_id,
	      w.work_date,
	      w.bldn_id,
	      w.workkind_id,
	      w.note,
	      w.print_flag,
	      w.private_note,
	      mw.workref_id
	FROM hidden_maintenance_works mw
	     INNER JOIN works w ON mw.workref_id = w.id;
COMMENT ON VIEW maintenance_works IS 'Список работ по содержанию';

CREATE VIEW streetlist AS
       SELECT s.* FROM streets s
       INNER JOIN villages v ON s.village_id = v.id
       ORDER BY v.md_id, v.name, s.name;

CREATE VIEW contractorlist AS
       SELECT * FROM contractors
       ORDER BY name;

CREATE VIEW xstreets AS
 SELECT s.id,
    v.id AS vid,
    md.id AS mid,
    concat_ws(' '::text, v.name, s.name, st.short_name) AS name,
    concat(v.site_name, s.site_name) AS site_name,
    concat_ws(' '::text, v.name, st.short_name, s.name) AS name1
   FROM streets s
     JOIN villages v ON v.id = s.village_id
     JOIN street_types st ON st.id = s.street_type
     JOIN municipal_districts md ON v.md_id = md.id;

CREATE VIEW current_subaccounts AS
       SELECT bs.*,
	      t.begin_date
       FROM bldn_subaccounts bs
       	    INNER JOIN terms t ON bs.term_id = t.id
       WHERE bs.term_id = (SELECT MAX(term_id) FROM bldn_subaccounts);

CREATE VIEW common_property_dictionary AS
  SELECT cpg.name AS group_name
	 , cpg.id AS group_id
	 , cpe.name AS element_name
	 , cpe.id AS element_id
	 , cpar.name AS parameter_name
	 , cpar.id AS parameter_id
    FROM common_property_element_parameter AS cpar
	 INNER JOIN common_property_element AS cpe ON cpar.element_id = cpe.id
	 INNER JOIN common_property_group AS cpg ON cpe.group_id = cpg.id;


CREATE FUNCTION get_bldn_works(bldnId INTEGER, gwtId INTEGER, wtId INTEGER, bTerm INTEGER, eTerm INTEGER, fSourceId INTEGER) RETURNS SETOF buildings_workslist AS
$$
BEGIN
	RETURN QUERY
	SELECT bw.*
	FROM buildings_workslist bw
	     INNER JOIN works w ON w.id = bw.id
	     INNER JOIN work_kinds wk ON wk.id = w.workkind_id
	WHERE w.bldn_id = bldnId
	      AND (term_id >= bTerm OR is_all_values(bTerm))
	      AND (term_id <= eTerm OR is_all_values(eTerm))
	      AND (bw.gwt_id = gwtId OR is_all_values(gwtId))
	      AND (wk.worktype_id = wtId OR is_all_values(wtId))
	      AND (w.finance_source = fSourceId OR is_all_values(fSourceId))
	ORDER BY bw.workdate DESC, bw.workkind_name;
END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION get_bldn_works IS 'Работы, проведенные на доме';

CREATE FUNCTION getBldnWorkYears(bldnid INTEGER, gwt INTEGER) RETURNS SETOF SMALLINT AS
$$
BEGIN
	RETURN QUERY SELECT DISTINCT(CAST(EXTRACT(YEAR FROM t.begin_date) AS SMALLINT)) FROM works w INNER JOIN terms t ON w.work_date = t.id WHERE w.bldn_id = bldnid AND w.gwt_id = gwt ORDER BY 1 DESC;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION getWorkYears() RETURNS SETOF SMALLINT AS
$$
BEGIN
	RETURN QUERY SELECT DISTINCT(CAST(EXTRACT(YEAR FROM t.begin_date) AS SMALLINT)) FROM works w INNER JOIN terms t ON w.work_date = t.id ORDER BY 1 DESC;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_service_modes(serviceId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF service_modes AS
$$
  SELECT * FROM service_modes WHERE service_id = serviceId ORDER BY id;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_service_modes IS 'Список режимов выбранной ЖКУ';

CREATE FUNCTION get_service_mode(modeId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS service_modes AS
$$
  SELECT * FROM service_modes WHERE id = modeId;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_service_mode IS 'Получение режима ЖКУ по коду';

CREATE FUNCTION create_service_mode(serviceId INTEGER, modeName VARCHAR, InUserId INTEGER, InPCName VARCHAR, OUT newid INTEGER) AS
$$
DECLARE
  _newsmId INTEGER;
BEGIN
  SELECT max(id)+1 INTO _newsmId FROM service_modes
   WHERE service_id = serviceId;
  
  IF _newsmId IS NULL THEN _newsmId := serviceId * 1000; END IF;
	
  INSERT INTO service_modes(id, service_id, mode_name)
  VALUES (_newsmId, serviceId, modeName);
  newid := _newsmId;

  INSERT INTO log_log(action_id, user_id, pc_name, log_action)
  SELECT 38, InUserId, InPCName, JSONB_AGG(ins_row)
    FROM service_modes AS ins_row
   WHERE id = newid;

END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_service_mode IS 'Добавление режима ЖКУ';

CREATE FUNCTION change_service_mode(itemId INTEGER, modeName VARCHAR, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    WITH updated_rows AS (
      UPDATE service_modes
	 SET mode_name = modeName
       WHERE id = itemId
	     RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 39, InUserId, InPCName, JSONB_AGG(ttt)
      FROM (SELECT JSONB_AGG(sm) AS prev, JSONB_AGG(updated_rows) AS upd
	      FROM service_modes AS sm, updated_rows
	     WHERE sm.id = itemId) AS ttt;
	
    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_service_mode IS 'Изменение режима ЖКУ';

CREATE FUNCTION delete_service_mode(itemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF itemId % 1000 = 0 THEN
	RAISE '%,%', get_error_number('delete_denied'), get_error_message('delete_denied');
    END IF;

    WITH deleted_rows AS (
      DELETE FROM service_modes
       WHERE id = itemId
      RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 40, InUserId, InPCName, JSONB_AGG(deleted_rows)
      FROM deleted_rows;
    
    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_service_mode IS 'Удаление режима ЖКУ';

CREATE FUNCTION get_expense_item(InItemId INTEGER, InUserId INTEGER, InPcName VARCHAR) RETURNS expense_items AS
$$
  SELECT * FROM expense_items WHERE id = InItemId;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_expense_item IS 'Статья расходов по коду';

CREATE FUNCTION get_expense_items(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF expense_items AS
$$
  SELECT * FROM expense_items ORDER BY short_name;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_expense_items IS 'Список статей расходов';

CREATE FUNCTION create_expense_item(InName1 VARCHAR, InName2 VARCHAR, InShortName VARCHAR(200), InGisGuid VARCHAR, InUkServiceId INTEGER, InGroupId INTEGER, InReportPriority INTEGER, InUseAsGroupName BOOLEAN, InUserId INTEGER, InPCName VARCHAR, OUT OutId INTEGER) AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
  
    INSERT INTO expense_items(name1, name2, short_name, gis_guid, uk_service_id, group_id, report_priority, use_as_group_name)
    VALUES (InName1, InName2, InShortName, InGisGuid, InUkServiceId, InGroupId, InReportPriority, InUseAsGroupName)
	   RETURNING id INTO OutId;

    INSERT INTO log_log (action_id, user_id, pc_name, action_description)
    SELECT
      10,
      InUserId,
      InPCName,
      CONCAT_WS(' :-: ', InName1, InName2, InShortName, InGisGuid, us.name, eg.name, InReportPriority, InUseAsGroupName)
      FROM uk_services us, expense_groups eg
     WHERE us.id = InUkServiceId AND eg.id = InGroupId;
      
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_expense_item IS 'Создание статьи расходов';

CREATE FUNCTION change_expense_item(InItemId INTEGER, InName1 VARCHAR, InName2 VARCHAR, InShortName VARCHAR, InGisGuid VARCHAR, InGroupId INTEGER, InReportPriority INTEGER, InUseAsGroupName BOOLEAN, InUkServiceId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
  
    WITH t AS (
      UPDATE expense_items
	 SET name1 = InName1,
	     name2 = InName2,
	     short_name = InShortName,
	     gis_guid = InGisGuid,
	     uk_service_id = InUkServiceId,
	     group_id = InGroupId,
	     report_priority = InReportPriority,
	     use_as_group_name = InUseAsGroupName
       WHERE id = InItemId
    )
	INSERT INTO log_log (action_id, user_id, pc_name, action_description)
    SELECT
      11,
      InUserId,
      InPCName,
      CONCAT_WS(' :-: ', 'Было', ei.name1, ei.name2, ei.short_name, ei.gis_guid, us1.name, eg1.name, ei.report_priority, ei.use_as_group_name, 'Стало', InName1, InName2, InShortName, InGisGuid, us.name, eg.name, InReportPriority, InUseAsGroupName)
      FROM uk_services us, expense_groups eg, expense_items ei
	   JOIN uk_services us1 ON us1.id = ei.uk_service_id
	   JOIN expense_groups eg1 ON eg1.id = ei.group_id
     WHERE us.id = InUkServiceId AND eg.id = InGroupId AND ei.id = InItemId;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_expense_item IS 'Изменение статьи расходов';

CREATE FUNCTION delete_expense_item(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_delete(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    WITH deleted_rows AS (
      DELETE FROM expense_items WHERE id = InItemId
      RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, action_description)
    SELECT
      12,
      InUserId,
      InPCName,
      CONCAT_WS(' :-: ', name1, name2, short_name, gis_guid, us.name, eg.name)
      FROM deleted_rows dr
	   JOIN uk_services us ON dr.uk_service_id = us.id
	   JOIN expense_groups eg ON dr.group_id = eg.id;

    RETURN;
  EXCEPTION WHEN foreign_key_violation THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_expense_item IS 'Удаление статьи расходов';

CREATE VIEW bldn_expenses AS
SELECT CASE ben.name_use WHEN 1 THEN ei.name1 ELSE ei.name2 END AS name,
       e.term_id,
       e.bldn_id,
       e.price,
       e.expense_plan_sum,
       e.expense_fact_sum,
       e.id,
       e.expense_item,
       ei.uk_service_id,
       ei.group_id,
       ei.report_priority
FROM expenses e
     INNER JOIN expense_items ei ON e.expense_item = ei.id
     INNER JOIN bldn_expense_names ben ON ben.bldn_id = e.bldn_id AND ben.expense_item = e.expense_item
		  ORDER BY e.term_id, ei.group_id, ei.uk_service_id, ei.report_priority;
COMMENT ON VIEW bldn_expenses IS 'Информация по структуре в разрезе домов с правильными названиями';


CREATE FUNCTION change_expense(itemId INTEGER, newprice NUMERIC(6,2), newplansum NUMERIC(14, 2), newfactsum NUMERIC(14,2)) RETURNS VOID AS
$$
BEGIN
	UPDATE expenses
	SET price = newprice,
	    expense_plan_sum = newplansum,
	    expense_fact_sum = newfactsum
	WHERE id = itemId;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION delete_expense(itemId INTEGER) RETURNS VOID AS
$$
BEGIN
	DELETE FROM expenses WHERE id = itemId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION add_expense(expenseId INTEGER, termId INTEGER, bldnId INTEGER, newprice NUMERIC(6, 2), newplansum NUMERIC(14, 2), newfactsum NUMERIC(14,2), InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
DECLARE
	bed INTEGER;
	bdt INTEGER;
BEGIN
	INSERT INTO expenses (id, expense_item, term_id, bldn_id, price, expense_plan_sum, expense_fact_sum)
	VALUES (DEFAULT, expenseId, termId, bldnId, newprice, newplansum, newfactsum);
	SELECT bldn_id INTO bed FROM bldn_expense_names
	WHERE bldn_id = bldnId AND expense_item = expenseId;
	IF bed IS NULL THEN
	   SELECT dogovor_type INTO bdt FROM buildings WHERE id = bldnId;
	   INSERT INTO bldn_expense_names VALUES (expenseId, bldnId, bdt);
	END if;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION delete_expenses_in_term(InTermId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DELETE FROM expenses WHERE term_id = InTermId;
$$ LANGUAGE SQL;

CREATE FUNCTION copy_expenses_from_term(oldTerm INTEGER, newTerm INTEGER) RETURNS VOID AS
$$
-- Пока не используется
BEGIN
	CREATE TEMPORARY TABLE _tmp ON COMMIT DROP AS
	SELECT * FROM expenses WHERE term_id = oldTerm;

	UPDATE _tmp SET term_id = newTerm;

	INSERT INTO expenses
	SELECT * FROM _tmp;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION delete_bldn_expenses(bldnId INTEGER, termId INTEGER) RETURNS VOID AS
$$
BEGIN
	DELETE FROM expenses
	WHERE bldn_id = bldnId
	      AND term_id = termId;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION bldn_list_use_expense_name(expId INTEGER, expname INTEGER) RETURNS TABLE (bldnId INTEGER, address TEXT) AS
$$
BEGIN
	RETURN QUERY
	SELECT ben.bldn_id, xs.name || ' д.' || b.bldn_no
	FROM buildings b
	     INNER JOIN xstreets xs ON b.street_id = xs.id
	     INNER JOIN bldn_expense_names ben ON b.id = ben.bldn_id
	WHERE ben.expense_item = expId
	      AND ben.name_use = expname
	ORDER BY 2;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION bldn_change_expense_name(bldnId INTEGER, expenseId INTEGER, expenseNameUse INTEGER) RETURNS VOID AS
$$
BEGIN
	UPDATE bldn_expense_names
	SET name_use = expenseNameUse
	WHERE bldn_id = bldnId
	      AND expense_item = expenseId;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION bldn_delete_expense_name(bldnId INTEGER, expenseId INTEGER) RETURNS VOID AS
$$
BEGIN
	DELETE FROM bldn_expense_names
	WHERE bldn_id = bldnId
	      AND expense_item = expenseId;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION bldn_add_expense_name(bldnId INTEGER, expenseId INTEGER, expenseName INTEGER DEFAULT 1) RETURNS VOID AS
$$
BEGIN
	INSERT INTO bldn_expense_names(expense_item, bldn_id, name_use)
	VALUES (bldnId, expenseId, expenseName);
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_bldn_expenses_in_term(bldnId INTEGER, paramId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF bldn_expenses AS
$$
  BEGIN
    RETURN QUERY
      SELECT * FROM bldn_expenses
      WHERE bldn_id = bldnId AND term_id = paramId;
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_bldn_expenses_in_term IS 'Структура дома в указанном месяце';

CREATE FUNCTION get_bldn_expense_history(bldnId INTEGER, paramId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF bldn_expenses AS
$$
  BEGIN
    RETURN QUERY
      SELECT be.* FROM bldn_expenses be
      INNER JOIN terms t ON be.term_id = t.id
      WHERE bldn_id = bldnId AND expense_item = paramId
      ORDER BY t.end_date DESC;
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_bldn_expense_history IS 'История изменения структуры по статье';

CREATE FUNCTION get_bldn_last_expenses(bldnId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF bldn_expenses AS
$$
  BEGIN
    RETURN QUERY
      SELECT * FROM bldn_expenses WHERE bldn_id = bldnId AND term_id = (SELECT MAX(term_id) FROM bldn_expenses) ORDER BY 1;
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_bldn_last_expenses IS 'Текущая структура в доме';

CREATE FUNCTION get_bldn_expense_terms(bldnId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF terms AS
$$
  BEGIN
    IF is_all_values(bldnId) THEN
      RETURN QUERY
      SELECT DISTINCT tr.*
      FROM terms AS tr
      INNER JOIN expenses AS ex ON ex.term_id = tr.id
      ORDER BY end_date DESC;
    ELSE 
      RETURN QUERY
	SELECT DISTINCT tr.*
	FROM terms AS tr
	INNER JOIN expenses AS ex ON ex.term_id = tr.id
	WHERE bldn_id = bldnId
	ORDER BY end_date DESC;
    END IF;
    RETURN;
END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_bldn_expense_terms IS 'Периоды в которых есть структура у дома';

CREATE FUNCTION add_bldn_service(bldnId INTEGER, serviceId INTEGER, modeId INTEGER, inputs INTEGER, possCounter BOOLEAN, newNote TEXT) RETURNS VOID AS
$$
DECLARE
	realPossCounter BOOLEAN;
	realInputs INTEGER;
BEGIN
	IF modeId % 1000 = 0
	THEN realPossCounter := False;
	     realInputs := 0;
	ELSE realPossCounter := possCounter;
	     realInputs := inputs;
	END IF;

	INSERT INTO bldn_services(bldn_id, service_id, mode_id, inputs_count, possible_counter, note) VALUES (bldnId, serviceId, modeId, realInputs, realPossCounter, newNote);
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION change_bldn_service(bldnId INTEGER, serviceId INTEGER, modeId INTEGER, inputs INTEGER, possCounter BOOLEAN, newNote TEXT) RETURNS VOID AS
$$
DECLARE
	realPossCounter BOOLEAN;
	realInputs INTEGER;
BEGIN
	IF modeId % 1000 = 0
	THEN realPossCounter := False;
	     realInputs := 0;
	ELSE realPossCounter := possCounter;
	     realInputs := inputs;
	END IF;

	UPDATE bldn_services
	SET mode_id = modeId,
	    inputs_count = realInputs,
	    possible_counter = realPossCounter,
	    note = newNote
	WHERE bldn_id = bldnId AND service_id = serviceID;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_bldn_services(bldnId INTEGER) RETURNS SETOF bldn_services AS
$$
BEGIN
	RETURN QUERY
	SELECT * FROM bldn_services WHERE bldn_id = bldnId
	ORDER BY service_id;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_service_in_bldn(bldnId INTEGER, serviceId INTEGER) RETURNS bldn_services AS
$$
	SELECT * FROM bldn_services WHERE bldn_id = bldnId AND service_id = serviceId;
$$ LANGUAGE SQL;

CREATE FUNCTION get_bldn_service_list(bldnId INTEGER) RETURNS SETOF bldn_services AS
$$
	SELECT * FROM bldn_services
	WHERE bldn_id = bldnId
	      AND mode_id % 1000 > 0;
$$ LANGUAGE SQL;

CREATE FUNCTION get_energo_classes() RETURNS SETOF energo_classes AS
$$
	SELECT * FROM energo_classes;
$$ LANGUAGE SQL;

CREATE FUNCTION get_roles() RETURNS SETOF roles AS
$$
BEGIN
	RETURN QUERY
	SELECT * FROM roles;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_user_roles(itemId INTEGER) RETURNS SETOF roles AS
$$
BEGIN
	RETURN QUERY
	SELECT r.* FROM roles r
	       INNER JOIN user_roles ur ON ur.role_id = r.id
	WHERE ur.user_id = itemId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_user_no_roles(itemId INTEGER) RETURNS SETOF roles AS
$$
BEGIN
	RETURN QUERY
	SELECT * FROM roles
	WHERE id NOT IN (SELECT role_id FROM user_roles WHERE user_id = itemId);
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION adm_has_admin_role(itemId INTEGER, OUT has_role BOOLEAN) AS
$$
	SELECT COUNT(*) > 0
	FROM user_roles ur
	     INNER JOIN users u ON ur.user_id = u.id
	WHERE ur.user_id = itemId
	     AND ur.role_id = 1
	     AND u.is_active;
$$ LANGUAGE SQL;

CREATE FUNCTION adm_add_user_role(itemId INTEGER, roleId INTEGER, userId INTEGER) RETURNS VOID AS
$$
BEGIN
	IF itemId = 0 THEN RAISE EXCEPTION '60010, не хватает прав'; END IF;	
	IF adm_has_admin_role(userId) THEN
	   	INSERT INTO user_roles(user_id, role_id)
		VALUES (itemId, roleId);
		RETURN;
	ELSE
		RAISE EXCEPTION '60010, не хватает прав';
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION adm_remove_user_role(itemId INTEGER, roleId INTEGER, userId INTEGER) RETURNS VOID AS
$$
BEGIN
	IF itemId = 0 THEN RAISE EXCEPTION '60010, не хватает прав'; END IF;
	IF adm_has_admin_role(userId) THEN
	   	DELETE FROM user_roles
		WHERE user_id = itemId AND role_id = roleId;
		RETURN;
	ELSE
		RAISE EXCEPTION '60010, не хватает прав';
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_user(itemId INTEGER) RETURNS users AS
$$
	SELECT * FROM users WHERE id = itemId;
$$ LANGUAGE SQL;

CREATE FUNCTION adm_get_users(userId INTEGER) RETURNS SETOF users AS
$$
BEGIN
	IF adm_has_admin_role(userId) THEN
	   	RETURN QUERY SELECT * FROM users ORDER BY name;
	ELSE
		RAISE EXCEPTION '60010, не хватает прав';
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION adm_create_user(newlogin VARCHAR(20), newname VARCHAR(200), newpwd TEXT, userId INTEGER, OUT newId INTEGER) AS
$$
BEGIN
	IF adm_has_admin_role(userId) THEN
	   	INSERT INTO users(login, name, password) VALUES (newlogin, newname, crypt(newpwd, gen_salt('md5')))
	   	RETURNING id INTO newId;
	ELSE
		RAISE EXCEPTION '60010, не хватает прав';
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION adm_change_username(itemId INTEGER, newname VARCHAR(200), userId INTEGER) RETURNS VOID AS
$$
BEGIN
	IF itemId = 0 THEN RAISE EXCEPTION '60010, не хватает прав'; END IF;
	IF adm_has_admin_role(userId) THEN
	   	UPDATE users SET name = newname WHERE id = itemId;
		RETURN;
	ELSE
		RAISE EXCEPTION '60010, не хватает прав';
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION adm_change_user_password(itemId INTEGER, newpwd TEXT, userId INTEGER) RETURNS VOID AS
$$
BEGIN
	IF itemId = 0 THEN RAISE EXCEPTION '60010, не хватает прав'; END IF;
	IF adm_has_admin_role(userId) THEN
	   	UPDATE users SET password = crypt(newpwd, gen_salt('md5')) WHERE id = itemId;
		RETURN;
	ELSE
		RAISE EXCEPTION '60010, не хватает прав';
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION adm_block_user(itemId INTEGER, userId INTEGER) RETURNS VOID AS
$$
BEGIN
	IF itemId = 0 THEN RAISE EXCEPTION '60010, не хватает прав'; END IF;
	IF adm_has_admin_role(userId) THEN
	   UPDATE users SET is_active = FALSE WHERE id = itemId;
	   RETURN;
	ELSE
		RAISE EXCEPTION '60010, не хватает прав';
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION adm_unblock_user(itemId INTEGER, userId INTEGER) RETURNS VOID AS
$$
BEGIN
	IF adm_has_admin_role(userId) THEN
	   UPDATE users SET is_active = TRUE WHERE id = itemId;
	   RETURN;
	ELSE
		RAISE EXCEPTION '60010, не хватает прав';
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION is_user_valid_password(userName VARCHAR(20), userPwd TEXT, OUT user_valid BOOLEAN) AS
$$
BEGIN
	SELECT ((password = crypt(userPwd, password)) AND is_active) FROM users WHERE login = userName INTO user_valid;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_user_info(userName VARCHAR(20), OUT userId INTEGER, OUT userFIO VARCHAR(200), OUT isAdmin BOOLEAN) AS
$$
BEGIN
	SELECT id, name, adm_has_admin_role(id) FROM users WHERE login = userName INTO userId, userFIO, isAdmin;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_user_info_by_id(userId INTEGER, OUT out_userName TEXT, OUT out_userFIO TEXT, OUT out_isAdmin BOOLEAN) AS
$$
BEGIN
	SELECT login, name, adm_has_admin_role(id) FROM users WHERE id = userId INTO out_userName, out_userFIO, out_isAdmin;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION adm_get_access_types(userId INTEGER) RETURNS SETOF access_types AS
$$
BEGIN
	IF adm_has_admin_role(userId) THEN
	   RETURN QUERY SELECT * FROM access_types;
	ELSE
		RAISE EXCEPTION SQLSTATE '60010';
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION adm_create_role(role_name VARCHAR(100), userId INTEGER) RETURNS VOID AS
$$
BEGIN
	IF adm_has_admin_role(userId) THEN
	   INSERT INTO roles(name) VALUES (role_name);
	ELSE
		RAISE EXCEPTION SQLSTATE '60010';
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION adm_role_has_access(roleId INTEGER, acsType INTEGER, userId INTEGER) RETURNS TABLE (id INTEGER, name VARCHAR(200)) AS
$$
DECLARE
	acsName ACCESS_TYPE;
BEGIN
	IF NOT adm_has_admin_role(userId) THEN RAISE EXCEPTION '60010'; END IF;

	SELECT acs_type INTO acsName FROM access_types act WHERE act.id = acsType;
	
	CASE acsName
	     WHEN 'gwt' THEN
	     	  RETURN QUERY
		  SELECT gwt.id, gwt.name FROM global_work_types gwt
		  	 INNER JOIN roles_access ra ON gwt.id = ra.acs_value
			 WHERE ra.role_id = roleId
			 AND ra.acs_id = acsType;
		  RETURN;
	END CASE;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION adm_role_has_no_access(roleId INTEGER, acsType INTEGER, userId INTEGER) RETURNS TABLE (id INTEGER, name VARCHAR(200)) AS
$$
DECLARE
	acsName ACCESS_TYPE;
BEGIN
	IF NOT adm_has_admin_role(userId) THEN RAISE EXCEPTION '60010'; END IF;

	SELECT acs_type INTO acsName FROM access_types act WHERE act.id = acsType;
	
	CASE acsName
	     WHEN 'gwt' THEN
	     	  RETURN QUERY
		  SELECT gwt.id, gwt.name FROM global_work_types gwt
		  WHERE gwt.id NOT IN
		  	(SELECT acs_value FROM roles_access ra
			WHERE ra.acs_id = acsType
			      AND ra.role_id = roleId);
	END CASE;
END;
$$ LANGUAGE plpgsql;
	
CREATE FUNCTION adm_add_role_access(roleId INTEGER, acsType INTEGER, acsVal INTEGER, userId INTEGER) RETURNS VOID AS
$$
BEGIN
	IF roleId = 1 THEN RAISE EXCEPTION 'Нельзя изменять права администратора'; END IF;
	IF adm_has_admin_role(userId) THEN
	   INSERT INTO roles_access(role_id, acs_id, acs_value) VALUES (roleId, acsType, acsVal);
	   RETURN;
	ELSE
		RAISE EXCEPTION SQLSTATE '60010';
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION adm_remove_role_access(roleId INTEGER, acsType INTEGER, acsVal INTEGER, userId INTEGER) RETURNS VOID AS
$$
BEGIN
	IF roleId = 1 THEN RAISE EXCEPTION 'Нельзя изменять права администратора'; END IF;
	IF adm_has_admin_role(userId) THEN
	   DELETE FROM roles_access
	   WHERE role_id = roleId
	   	 AND acs_id = acsType
		 AND ((acs_value = acsVal) OR ((acs_value IS NULL) AND (acsVal IS NULL)));
	ELSE
		RAISE EXCEPTION SQLSTATE '60010';
	END IF;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION has_access(userId INTEGER, acsType ACCESS_TYPE, acsVal INTEGER) RETURNS BOOLEAN AS
$$
DECLARE
	return_value BOOLEAN;
BEGIN
	IF adm_has_admin_role(userId) AND acsType <> 'gwt' THEN RETURN TRUE; END IF;
	
	SELECT COUNT(*) > 0
	FROM roles_access ra
	     INNER JOIN user_roles ur ON ra.role_id = ur.role_id
	     INNER JOIN access_types act ON ra.acs_id = act.id
	     WHERE act.acs_type = acsType
	     AND ur.user_id = userId
	     AND ((ra.acs_value = acsVal) OR (ra.acs_value IS NULL AND acsVal IS NULL))
	INTO return_value;
	RETURN return_value;
END;
$$ LANGUAGE plpgsql;


CREATE FUNCTION add_access_right_trigger() RETURNS TRIGGER AS
$$
BEGIN
  INSERT INTO roles_access_rights(access_id, role_id)
  SELECT NEW.id, r.id FROM roles r;

  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER access_rights__insert__trigger AFTER INSERT ON access_rights FOR EACH ROW EXECUTE PROCEDURE add_access_right_trigger();

CREATE FUNCTION delete_access_right_trigger() RETURNS TRIGGER AS
$$
  BEGIN
    DELETE FROM roles_access_rights
     WHERE access_id = OLD.id;

    RETURN OLD;
  END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER access_rights__delete__trigger BEFORE DELETE ON access_rights FOR EACH ROW EXECUTE PROCEDURE delete_access_right_trigger();


CREATE FUNCTION add_role_trigger() RETURNS TRIGGER AS
$$
BEGIN
  INSERT INTO roles_access_rights(access_id, role_id)
  SELECT ar.id, NEW.id FROM access_rights ar;

  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER roles__insert__trigger AFTER INSERT ON roles FOR EACH ROW EXECUTE PROCEDURE add_role_trigger();

CREATE FUNCTION delete_role_trigger() RETURNS TRIGGER AS
$$
  BEGIN
    DELETE FROM roles_access_rights
     WHERE role_id = OLD.id;

    RETURN OLD;
  END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER roles__delete__trigger BEFORE DELETE ON roles FOR EACH ROW EXECUTE PROCEDURE delete_role_trigger();

CREATE FUNCTION add_gwt_trigger() RETURNS TRIGGER AS
$$
BEGIN
  INSERT INTO access_rights(id, name)
  SELECT CAST(c.value AS INTEGER) + NEW.id, 'Работы ' || NEW.name
	   FROM constants c WHERE c.name = 'gwt_work_access_prefix';

  INSERT INTO access_rights(id, name)
  SELECT CAST(c.value AS INTEGER) + NEW.id, 'Планы ' || NEW.name
	   FROM constants c WHERE c.name = 'gwt_planwork_access_prefix';

  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER gwt__insert__trigger AFTER INSERT ON global_work_types FOR EACH ROW EXECUTE PROCEDURE add_gwt_trigger();

CREATE FUNCTION delete_gwt_trigger() RETURNS TRIGGER AS
$$
  DECLARE
    _gwt_pref INTEGER;
    _plan_pref INTEGER;
  BEGIN

    SELECT CAST(c.value AS INTEGER) INTO _gwt_pref FROM constants c WHERE c.name = 'gwt_work_access_prefix';
    SELECT CAST(c.value AS INTEGER) INTO _plan_pref FROM constants c WHERE c.name = 'gwt_planwork_access_prefix';

    DELETE FROM access_rights
     WHERE id = _gwt_pref + OLD.id;

    DELETE FROM access_rights
     WHERE id = _plan_pref + OLD.id;

    RETURN OLD;
  END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER gwt__delete__trigger BEFORE DELETE ON global_work_types FOR EACH ROW EXECUTE PROCEDURE delete_gwt_trigger();

CREATE FUNCTION user_has_right_read(InUserId INTEGER, InAccessId INTEGER, OUT OutHasRight BOOLEAN) AS
$$
  SELECT u.is_active AND EXISTS (
    SELECT role_id FROM user_roles WHERE user_id = u.id
    INTERSECT
    SELECT role_id FROM roles_access_rights WHERE access_read AND access_id = InAccessId
  ) FROM users u WHERE u.id = InUserId;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION user_has_right_change(InUserId INTEGER, InAccessId INTEGER, OUT OutHasRight BOOLEAN) AS
$$
  SELECT u.is_active AND EXISTS (
    SELECT role_id FROM user_roles WHERE user_id = u.id
    INTERSECT
    SELECT role_id FROM roles_access_rights WHERE access_change AND access_id = InAccessId
  ) FROM users u WHERE u.id = InUserId;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION user_has_right_delete(InUserId INTEGER, InAccessId INTEGER, OUT OutHasRight BOOLEAN) AS
$$
  SELECT u.is_active AND EXISTS (
    SELECT role_id FROM user_roles WHERE user_id = u.id
    INTERSECT
    SELECT role_id FROM roles_access_rights WHERE access_delete AND access_id = InAccessId
  ) FROM users u WHERE u.id = InUserId;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION add_tmp_counter(InBldnId INTEGER, InName VARCHAR, InUserId INTEGER, InPCName VARCHAR, OUT OutId INTEGER) AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_counter_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    INSERT INTO tmp_counters(id, bldn_id, name)
    VALUES (DEFAULT, InBldnId, InName)
	   RETURNING id INTO OutId;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION change_tmp_counter(InItemId INTEGER, InName VARCHAR, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_counter_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    UPDATE tmp_counters
       SET name = InName
     WHERE id = InItemId;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION delete_tmp_counter(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_delete(InUserId, rights_get_counter_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    DELETE FROM tmp_counters
     WHERE id = InItemId;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_all_tmp_counters(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF certificate_tmp_counters AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_counter_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    RETURN QUERY
    SELECT * FROM certificate_tmp_counters ORDER BY end_date DESC;
  END;
$$ LANGUAGE plpgsql STABLE;

CREATE FUNCTION get_bldn_tmp_counters(InBldnId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF certificate_tmp_counters AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_counter_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    RETURN QUERY
      SELECT *
      FROM certificate_tmp_counters
      WHERE bldn_id = InBldnId
      ORDER BY end_date DESC;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_tmp_counter(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS tmp_counters AS
$$
  DECLARE _tmp_counter tmp_counters%ROWTYPE;
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_counter_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
    
    SELECT * INTO _tmp_counter FROM tmp_counters WHERE id = InItemId;
    RETURN _tmp_counter;
  END;
$$ LANGUAGE plpgsql STABLE;

CREATE FUNCTION add_counter_certificate(InCounterId INTEGER, InDate DATE, InNote TEXT, InValidite INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_certificate_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    INSERT INTO certificates(counter_id, certificate_date, certificate_validite, note)
    VALUES (InCounterId, InDate, InValidite, InNote);

    RETURN;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_counter_certificates(InCounterId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF certificates AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_certificate_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
    
    RETURN QUERY
      SELECT * FROM certificates WHERE counter_id = InCounterId ORDER BY (certificate_date + MAKE_INTERVAL(years => certificate_validite)) DESC;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION delete_counter_certificate(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_delete(InUserId, rights_get_certificate_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    DELETE FROM certificates WHERE id = InItemId;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION add_signature(InBeginTerm INTEGER, InBldnId INTEGER, NotIsChairman BOOLEAN, InOwner VARCHAR, InSign BYTEA, InUserId INTEGER, InPCName VARCHAR) RETURNS void AS
$$
  DECLARE
    _chairman_name VARCHAR;
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_owners_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    IF NotIsChairman THEN
      _chairman_name := InOwner;
    ELSE
      SELECT * INTO _chairman_name
      FROM get_bldn_chairman_in_term(InBldnId, InBeginTerm, InUserId, InPCName);
    END IF;
    
    INSERT INTO chairman_signature(begin_term, bldn_id, sign, signature_owner)
    VALUES (InBeginTerm, InBldnId, InSign, _chairman_name);

    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 107, InUserId, InPCName, JSONB_BUILD_OBJECT('bldn_id', InBldnId, 'begin_term', InBeginTerm, 'owner_name', _chairman_name,
						       'not_is_chairman', NotIsChairman);

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION add_signature IS 'Добавление подписи';

CREATE FUNCTION get_signature_in_term(InBldnId INTEGER, InTermId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS chairman_signature AS
$$
  DECLARE
    _out_row chairman_signature%ROWTYPE;
  BEGIN
    SELECT * INTO _out_row
      FROM chairman_signature
     WHERE bldn_id = InBldnId
       AND begin_term <= InTermId
     ORDER BY begin_term DESC
     LIMIT 1;

    IF _out_row IS NULL THEN
      SELECT InTermId, InBldnId, NULL::BYTEA, get_bldn_chairman_in_term(InBldnId, InTermId, InUserId, InPCName) INTO _out_row;
    END IF;

    RETURN _out_row;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_signature_in_term IS 'Получение подписи коменданта в периоде';

CREATE FUNCTION get_bldn_chairman_in_term(InBldnId BIGINT, InTermId INTEGER, InUserId INTEGER, InPCName VARCHAR, OUT OutName VARCHAR) AS
$$
  BEGIN
    IF is_not_value(InTermId) THEN
      SELECT MAX(term_id) INTO InTermId FROM flats;
    END IF;

    SELECT owner_name INTO OutName
      FROM owners AS ow
	   INNER JOIN flat_shares AS fls ON ow.share_id = fls.id
	   INNER JOIN flats AS f USING (flat_id, term_id)
     WHERE f.bldn_id = InBldnId
       AND f.term_id = InTermId
       AND ow.is_chairman;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_bldn_chairman_in_term IS 'ФИО коменданта в указанном периоде';

CREATE FUNCTION delete_chairman_signature(InBldnId INTEGER, InTermId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_owners_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    WITH deleted_rows AS (
      DELETE FROM chairman_signature
       WHERE bldn_id = InBldnId
	 AND begin_term = InTermId
      RETURNING *
    )
    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 108, InUserId, InPCName, JSONB_AGG(deleted_rows)
      FROM deleted_rows;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_chairman_signature IS 'Удаление подписи коменданта в указанном периоде';


CREATE FUNCTION get_chairmans_signature_in_bldn(InBldnId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF chairman_signature AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_owners_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    RETURN QUERY
    SELECT *
      FROM chairman_signature
      WHERE bldn_id = InBldnId
      ORDER BY begin_term DESC;

    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_chairmans_signature_in_bldn IS 'Список подписей комендантов в доме';  

CREATE FUNCTION get_bldn_id_no_list(InStreetId INTEGER, InVillageId INTEGER, InMdId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF bldn_id_no_list AS
$$
  BEGIN
    RETURN QUERY
    SELECT * FROM bldn_id_no_list
    WHERE (street_id = InStreetId OR is_all_values(InStreetId))
     AND (village_id = InVillageId OR is_all_values(InVillageId))
     AND (md_id = InMdId OR is_all_values(InMdId));
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_bldn_id_no_list IS 'Список домов с адресом';

CREATE FUNCTION get_managed_bldn_id_no_list(InStreetId INTEGER, InVillageId INTEGER, InMdId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF bldn_id_no_list AS 
$$
  BEGIN
    RETURN QUERY
    SELECT blist.*
      FROM bldn_id_no_list AS blist
	   INNER JOIN managed_buildings AS mlist ON mlist.id = blist.bid
     WHERE (blist.street_id = InStreetId OR is_all_values(InStreetId))
      AND (village_id = InVillageId OR is_all_values(InVillageId))
      AND (md_id = InMdId OR is_all_values(InMdId));
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_managed_bldn_id_no_list IS 'Список обслуживаемых домов с адресом';

CREATE FUNCTION add_building() RETURNS TRIGGER AS
$$
  BEGIN
    INSERT INTO buildings_land_info(bldn_id)
    VALUES (NEW.id);
    INSERT INTO buildings_tech_info(bldn_id)
    VALUES (NEW.id);
  
    INSERT INTO building_common_property_elements(bldn_id, element_id)
    SELECT NEW.id, cpe.id
      FROM common_property_element AS cpe;		 
    
    INSERT INTO bldn_services (bldn_id, service_id, mode_id)
    SELECT NEW.id, ss.id, ss.id * 1000
      FROM services AS ss;

    RETURN NEW;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION add_building IS 'Процедуры необходимые при создании дома';

CREATE TRIGGER buildings__insert__index AFTER INSERT ON buildings FOR EACH ROW EXECUTE PROCEDURE add_building();

CREATE FUNCTION drop_building() RETURNS TRIGGER AS
$BODY$
BEGIN
	DELETE FROM buildings_land_info
	       WHERE bldn_id=OLD.id;
	DELETE FROM buildings_tech_info
	       WHERE bldn_id = OLD.id;

	 RETURN OLD;
END;
$BODY$ LANGUAGE plpgsql;

CREATE TRIGGER buildings__delete__trigger BEFORE DELETE ON buildings FOR EACH ROW EXECUTE PROCEDURE drop_building();

CREATE FUNCTION report_1(mdid INTEGER, mcid INTEGER, contid INTEGER, dogid INTEGER, gwtid INTEGER, wtid INTEGER, wkid INTEGER, pf BOOLEAN, bdate INTEGER, edate INTEGER) RETURNS TABLE (V01 VARCHAR(20), V02 VARCHAR(100), V03 VARCHAR(300), V04 INTEGER, V05 VARCHAR(300), V06 TEXT, V07 VARCHAR(150), V08 VARCHAR(20), V09 VARCHAR(50), V10 NUMERIC(12,2)) AS
$$
BEGIN
	CREATE TEMPORARY TABLE _tmp(report_name VARCHAR(20), cont_name VARCHAR(100), address VARCHAR(300), bldn_id INTEGER, worktype INTEGER, wt_name VARCHAR(300), wk_name VARCHAR(200), note TEXT, dogovor VARCHAR(150), si VARCHAR(20), volume VARCHAR(50), work_sum NUMERIC(12,2), mancompid INTEGER, mundisid INTEGER, contrid INTEGER, dogovorid INTEGER, globwtid INTEGER, workkid INTEGER, prinf BOOLEAN) ON COMMIT DROP;

	INSERT INTO _tmp
	SELECT mc.report_name, 
		c.name, 
		xs.name || ' д.' || b.bldn_no,
		b.id,
		wt.id, 
		wt.name, 
		wk.name,
		COALESCE(' (' || w.note || ')', ''), 
		w.dogovor, 
		w.si, 
		w.volume, 
		w.work_sum,
		mc.id,
		md.id,
		c.id,
		d.id,
		gwt.id,
		wk.id,
		w.print_flag
	FROM works w
	INNER JOIN buildings b ON w.bldn_id = b.id
	INNER JOIN xstreets xs ON b.street_id = xs.id
	INNER JOIN contractors c ON w.contractor_id = c.id
	INNER JOIN work_kinds wk ON w.workkind_id = wk.id
	INNER JOIN work_types wt ON wk.worktype_id = wt.id
	INNER JOIN global_work_types gwt ON w.gwt_id = gwt.id
	INNER JOIN management_companies mc ON w.mc_id = mc.id
	INNER JOIN municipal_districts md ON md.id = xs.mid
	INNER JOIN dogovors d ON b.dogovor_type = d.id
	WHERE w.work_date BETWEEN bdate AND edate;

	IF mdid <> -1002 THEN DELETE FROM _tmp WHERE mundisid <> mdid; END IF;
	IF mcid <> -1002 THEN DELETE FROM _tmp WHERE mancompid <> mcid; END IF;
	IF contid <> -1002 THEN DELETE FROM _tmp WHERE contrid <> contid; END IF;
	IF dogid <> -1002 THEN DELETE FROM _tmp WHERE dogovorid <> dogid; END IF;
	IF gwtid <> -1002 THEN DELETE FROM _tmp WHERE globwtid <> gwtid; END IF;
	IF wtid <> -1002 THEN DELETE FROM _tmp WHERE worktype <> wtid; END IF;
	IF wkid <> -1002 THEN DELETE FROM _tmp WHERE workkid <> wkid; END IF;
	IF NOT pf IS NULL THEN DELETE FROM _tmp WHERE prinf <> pf; END IF;

	CREATE TEMPORARY TABLE _tmp1 ON COMMIT DROP AS
	SELECT report_name, cont_name, address, bldn_id, worktype, wt_name, wk_name, note, dogovor, si, volume, work_sum
FROM _tmp
	UNION ALL
	SELECT distinct NULL, NULL, NULL, 0, worktype, NULL, wk_name, NULL, NULL, NULL, NULL, 0
	FROM _tmp;

	RETURN QUERY
	SELECT report_name, cont_name, address, bldn_id, wt_name, CASE WHEN report_name IS NULL THEN '' ELSE wk_name || note END, dogovor, si, volume, work_sum
	FROM _tmp1
	ORDER BY worktype, wk_name, CASE WHEN address IS NULL THEN 'яяяяяяяяяяяяяяяяяяяяя' ELSE address END;

END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION report_2(mcid INTEGER, dogovor INTEGER, mdid INTEGER, contid INTEGER, InUserId INTEGER, InPcName VARCHAR)
RETURNS TABLE(c_id INTEGER, c_address TEXT, c_dogname VARCHAR(50), c_heating INTEGER, c_hotwater INTEGER, c_gas INTEGER, c_floormin INTEGER, c_floormax INTEGER, c_wmname VARCHAR(100), c_builtyear INTEGER, c_commyear INTEGER, c_totalsq NUMERIC(12, 2), c_stairssq REAL, c_corrsq REAL, c_othersq REAL, c_mop REAL, c_entrances INTEGER, c_stairs INTEGER, c_vaults INTEGER, c_vaultssq REAL, c_atticsq REAL, c_structvol REAL, c_depr REAL, c_outr BOOLEAN, c_odpuhw BOOLEAN, c_odpuheating BOOLEAN, c_odpucommon BOOLEAN, c_odpucw BOOLEAN, c_odpuee BOOLEAN, c_hasdoorphone BOOLEAN, c_squarebanister REAL, c_squaredoors REAL, c_squarewindowsills REAL, c_squaredoorhandles REAL, c_squaremailboxes REAL, c_squareradiators REAL, c_dogid INTEGER, c_mcid INTEGER, c_mdid INTEGER, c_contid INTEGER, c_hasthermoregulator BOOLEAN,  c_hasdoorcloser BOOLEAN) AS
$$
  BEGIN
    RETURN QUERY
      WITH t AS (
	SELECT
	  bldn_id
	  , SUM(square) AS total_square
	  FROM flats
	 WHERE term_id = (SELECT max(term_id) FROM flats)
	 GROUP by bldn_id
      )
      SELECT b.id,
      xs.name || ' д. ' || b.bldn_no,
      d.short_name,
      b.heating,
      b.hot_water,
      b.gas,
      bti.floor_min,
      bti.floor_max,
      wm.name,
      bti.built_year,
      bti.commissioning_year,
      COALESCE(t.total_square, 0),
      bti.stairs_square,
      bti.corridor_square,
      bti.other_square, 
      bti.stairs_square + bti.corridor_square + coalesce(bti.other_square, 0),
      bti.entrances,
      bti.stairs,
      bti.vaults,
      bti.vaults_square,
      bti.attic_square,
      bti.structural_volume,
      bti.depreciation,
      b.out_report,
      bti.has_odpu_hotwater,
      bti.has_odpu_heating,
      bti.has_odpu_common,
      bti.has_odpu_coldwater,
      bti.has_odpu_electro,
      bti.has_doorphone,
      bti.square_banisters,
      bti.square_doors,
      bti.square_windowsills,
      bti.square_doorhandles,
      bti.square_mailboxes,
      bti.square_radiators,
      d.id,
      b.mc_id,
      xs.mid,
      b.contractor_id,
      bti.has_thermoregulator,
      bti.has_doorcloser
      FROM buildings b
      INNER JOIN xstreets xs ON b.street_id = xs.id
      INNER JOIN buildings_tech_info bti ON b.id = bti.bldn_id
      INNER JOIN wall_materials wm ON bti.wallmater_id = wm.id
      INNER JOIN dogovors d ON b.dogovor_type = d.id
      LEFT JOIN t ON b.id = t.bldn_id
      WHERE b.out_report
      AND (is_all_values(mcid) OR mcid = b.mc_id)
      AND (is_all_values(dogovor) OR dogovor = d.id)
      AND (is_all_values(mdid) OR mdid = xs.mid)
      AND (is_all_values(contid) OR contid = b.contractor_id)
      ORDER BY xs.mid, xs.vid, 1;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION report_2 IS 'Отчет по технической информации';

CREATE FUNCTION report_3(mcid INTEGER, dogid INTEGER, mdid INTEGER, contid INTEGER)
RETURNS TABLE (c_bid INTEGER, c_address TEXT, c_dogname VARCHAR(20), c_cadastral VARCHAR(20), c_invarea REAL, c_usearea REAL, c_survarea REAL, c_builtuparea REAL, c_undevarea REAL, c_hardcoat REAL, c_driveways REAL, c_sidewalks REAL, c_otherhard REAL, c_saf BOOLEAN, c_fences BOOLEAN, c_benches INTEGER, c_mcid INTEGER, c_dogid INTEGER, c_mdid INTEGER, c_contid INTEGER) AS
$$
BEGIN
	CREATE TEMPORARY TABLE _tmp ON COMMIT DROP AS
	SELECT b.id AS b_id,
       	       xs.name || ' д. ' || b.bldn_no AS address,
       	       d.short_name AS dog_name,
       	       bli.cadastral_no AS cadastral_no,
       	       bli.inventory_area AS inv_area,
       	       bli.use_area AS use_area,
       	       bli.survey_area AS surv_area,
       	       bli.builtup_area AS built_area,
       	       bli.undeveloped_area AS undev_area,
       	       bli.hard_coatings AS hard_coatings,
       	       bli.drive_ways_hard AS drive_ways,
       	       bli.side_walks_hard AS side_walks,
       	       bli.others_hard AS others_hars,
	       bli.saf AS saf,
	       bli.fences AS fences,
	       bli.benches AS benches,
	       b.mc_id AS mc_id,
	       b.dogovor_type AS dog_id,
	       xs.mid as md_id,
	       b.contractor_id as cont_id
	FROM buildings b
	INNER JOIN xstreets xs ON b.street_id = xs.id
	INNER JOIN buildings_land_info bli ON b.id = bli.bldn_id
	INNER JOIN dogovors d ON b.dogovor_type = d.id
	WHERE b.out_report
	ORDER BY xs.mid, xs.vid, address;

	IF mcid <> -1002 THEN DELETE FROM _tmp WHERE mc_id <> mcid; END IF;
	IF dogid <> -1002 THEN DELETE FROM _tmp WHERE dog_id <> dogid; END IF;
	IF mdid <> -1002 THEN DELETE FROM _tmp WHERE md_id <> mdid; END IF;
	IF contid <> -1002 THEN DELETE FROM _tmp WHERE cont_id <> contid; END IF;

	RETURN QUERY SELECT * FROM _tmp;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION report_4(mcid INTEGER, mdid INTEGER, contid INTEGER, gwt INTEGER, wt INTEGER, wk INTEGER, pstat INTEGER, dogid INTEGER, bdate DATE, edate DATE) RETURNS TABLE (c_mc_name VARCHAR(100), c_cont_name VARCHAR(200), c_address TEXT, c_bid INTEGER, c_gwtid INTEGER, c_gwtname VARCHAR(100), c_wtid INTEGER, c_wtname VARCHAR(200), c_wkname TEXT, c_wdate DATE, c_statname VARCHAR(100), c_empname VARCHAR(300), c_wsum NUMERIC(12,2), c_ssum NUMERIC(12,2), b_date DATE, e_date DATE) AS
$$
BEGIN
	CREATE TEMPORARY TABLE _tmp ON COMMIT DROP AS
	SELECT mc.report_name AS mc_name, 
       	       c.name AS cont_name, 
       	       xs.name || ' д.' || b.bldn_no AS address,
       	       b.id AS bid,
       	       gwt.id AS gwtid,
       	       gwt.name AS gwtname,
       	       wt.id AS wtid,
       	       wt.name AS wtname, 
       	       wk.name || COALESCE(' (' || w.note || ')', '') AS wkname, 
       	       w.work_date AS wdate,
       	       ps.name AS statname,
       	       w.employee AS empname,
       	       w.work_sum AS work_sum,
	       w.smeta_sum AS smeta_sum,
	       mc.id AS mc_id,
	       c.id AS cont_id,
	       wk.id AS wk_id,
	       ps.id AS ps_id,
               b.dogovor_type,
	       w.begin_date,
	       w.end_date
	 FROM plan_works w
	 INNER JOIN buildings b ON w.bldn_id = b.id
	 INNER JOIN xstreets xs ON b.street_id = xs.id
	 INNER JOIN contractors c ON w.contractor_id = c.id
	 INNER JOIN work_kinds wk ON w.workkind_id = wk.id
	 INNER JOIN work_types wt ON wk.worktype_id = wt.id
	 INNER JOIN global_work_types gwt ON w.gwt_id = gwt.id
	 INNER JOIN management_companies mc ON w.mc_id = mc.id
	 INNER JOIN municipal_districts md ON md.id = xs.mid
	 INNER JOIN dogovors d ON b.dogovor_type = d.id
	 INNER JOIN plan_work_statuses ps ON w.work_status = ps.id
	 WHERE w.work_date BETWEEN bdate AND edate
	 ORDER BY w.work_date, gwt.id, wt.name, wk.name;

	 IF mcid <> -1002 THEN DELETE FROM _tmp WHERE mc_id <> mcid; END IF;
	 IF mdid <> -1002 THEN DELETE FROM _tmp WHERE md_id <> mdid; END IF;
	 IF contid <> -1002 THEN DELETE FROM _tmp WHERE cont_id <> contid; END IF;
	 IF gwt <> -1002 THEN DELETE FROM _tmp WHERE gwtid <> gwt; END IF;
	 IF wt <> -1002 THEN DELETE FROM _tmp WHERE wtid <> wt; END IF;
	 IF wk <> -1002 THEN DELETE FROM _tmp WHERE wk_id <> wk; END IF;
	 IF pstat <> -1002 THEN DELETE FROM _tmp WHERE ps_id <> pstat; END IF;
         IF dogid <> -1002 THEN DELETE FROM _tmp WHERE dogovor_type <> dogid; END IF;

	 RETURN QUERY
	 SELECT mc_name, cont_name, address, bid, gwtid, gwtname, wtid, wtname, wkname, wdate, statname, empname, work_sum, smeta_sum, begin_date, end_date FROM _tmp;

END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION report_5(bdate INTEGER, edate INTEGER, contid INTEGER, gwt INTEGER)
RETURNS TABLE(out_bldnid INTEGER, out_contractorname VARCHAR(100), out_address TEXT, out_worktype VARCHAR(100), out_workname TEXT, out_worksum NUMERIC(15,2), out_workdate DATE, out_dogovor VARCHAR(200), out_volume TEXT, out_contid INTEGER, out_gwtid INTEGER, out_mdid INTEGER) AS
$$
BEGIN
	CREATE TEMPORARY TABLE _tmp ON COMMIT DROP AS
	SELECT b.id AS bid,
       	       c.name AS contractor_name,
       	       xs.name || ' д.' || b.bldn_no AS address,
       	       wt.name AS worktype_name,
       	       wk.name || coalesce(' (' || w.note || ')', '') AS work_name,
       	       w.work_sum AS work_sum,
       	       t.begin_date AS work_date,
       	       w.dogovor AS dogovor,
       	       w.volume || ' ' || w.si AS work_volume,
       	       c.id AS contractor_id,
	       w.gwt_id AS gwt_id,
	       xs.mid AS md_id
	FROM works w
	     INNER JOIN work_kinds wk ON w.workkind_id = wk.id
	     INNER JOIN work_types wt ON wk.worktype_id = wt.id
	     INNER JOIN buildings b ON w.bldn_id = b.id
	     INNER JOIN xstreets xs ON b.street_id = xs.id
	     INNER JOIN contractors c ON w.contractor_id = c.id
	     INNER JOIN terms t ON w.work_date = t.id
	     WHERE  w.print_flag
	     AND t.id BETWEEN bdate AND edate
	     AND (c.id = contid OR is_all_values(contid))
	     AND (w.gwt_id = gwt OR is_all_values(gwt))
	     ORDER BY c.id, wt.id, xs.mid, 3;
	     
	RETURN QUERY SELECT * FROM _tmp ORDER BY contractor_name, worktype_name, md_id, address;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION report_6(InBDate INTEGER, InEDate INTEGER, InContId INTEGER, InDogovor INTEGER, InBYear INTEGER, InEYear INTEGER, InUserId INTEGER, InPCName TEXT)
RETURNS TABLE(V01 INTEGER, V02 VARCHAR(20), V03 VARCHAR(100), V04 TEXT, V05 NUMERIC, V06 NUMERIC(15,2), V07 NUMERIC(15,2), V08 NUMERIC(15,2), V09 NUMERIC(15,2), V10 NUMERIC(15,2), V11 BIGINT, V12 NUMERIC(15,2), V13 INTEGER, V14 INTEGER) AS
$$
  BEGIN
    RETURN QUERY
      WITH _lists AS (
	SELECT DISTINCT bldn_id, contractor_id
	  FROM accrueds
	 WHERE acc_date BETWEEN InBYear AND InEYear
	   AND (contractor_id = InContId OR is_all_values(InContId))
	 UNION
	SELECT DISTINCT bldn_id, contractor_id
	  FROM works
	 WHERE gwt_id = 1
	   AND work_date BETWEEN InBDate AND InEDate
	   AND (contractor_id = InContId OR is_all_values(InContId))
      ), _a AS (
	SELECT bldn_id,
	       SUM(acc_sum) AS acc_sum,
	       COUNT(acc_sum) AS count_sum,
	       contractor_id
	  FROM accrueds
	 WHERE acc_date BETWEEN InBYear AND InEYear
	   AND (contractor_id = InContId OR is_all_values(InContId))
	 GROUP BY bldn_id, contractor_id
      ), _w AS (
	SELECT bldn_id,
	       SUM(work_sum) AS work_sum,
	       contractor_id
	  FROM works
	 WHERE gwt_id = 1
	   AND work_date BETWEEN InBDate AND InEDate
	   AND print_flag
	   AND (contractor_id = InContId OR is_all_values(InContId))
	 GROUP BY bldn_id, contractor_id
      ), _a2 AS (
	SELECT bldn_id,
	       SUM(acc_sum) AS acc_sum,
	       contractor_id
	  FROM accrueds
	 WHERE acc_date BETWEEN InBDate AND InEDate
	   AND (contractor_id = InContId OR is_all_values(InContId))
	 GROUP BY bldn_id, contractor_id
      ), t AS (
	SELECT
	  bldn_id
	  , SUM(square) AS total_square
	  FROM flats
	 WHERE term_id = (SELECT max(term_id) FROM flats)
	 GROUP by bldn_id
      ), _out_table AS (	
	SELECT
	  b.id, 
	  mc.report_name, 
	  c.name,
	  xs.name || ' д.' || b.bldn_no, 
	  t.total_square, 
	  COALESCE((SELECT work_sum FROM works WHERE bldn_id = b.id AND workkind_id = 1 AND work_date BETWEEN InBDate AND InEDate AND contractor_id = l.contractor_id ORDER BY work_date DESC LIMIT 1), 0) AS month_avr, 
	  COALESCE((SELECT acc_sum FROM accrueds WHERE bldn_id = b.id AND acc_date BETWEEN InBDate AND InEDate AND contractor_id = l.contractor_id ORDER BY acc_date DESC LIMIT 1), 0) AS month_acc, 
	  COALESCE(_a2.acc_sum, 0) AS acc_period, 
	  COALESCE(_a.acc_sum, 0),
	  COALESCE((SELECT a2.acc_sum FROM accrueds a2 WHERE a2.bldn_id = b.id AND a2.acc_date = (SELECT MAX(acc_date) FROM accrueds a1 WHERE a1.acc_date BETWEEN InBDate AND InEDate AND contractor_id = l.contractor_id)),0), 
	  COALESCE(_a.count_sum,0), 
	  COALESCE(_w.work_sum,0) AS work_sum_period,
	  c.id,
	  xs.mid
	  FROM _lists l
	       INNER JOIN buildings b ON b.id = l.bldn_id
	       INNER JOIN xstreets xs ON b.street_id = xs.id
	       INNER JOIN management_companies mc ON b.mc_id = mc.id
	       INNER JOIN contractors c ON l.contractor_id = c.id
	       LEFT JOIN t ON t.bldn_id = b.id
	       LEFT JOIN _a ON (b.id = _a.bldn_id AND l.contractor_id = _a.contractor_id)
	       LEFT JOIN _w ON (b.id = _w.bldn_id AND l.contractor_id= _w.contractor_id)
	       LEFT JOIN _a2 ON (b.id = _a2.bldn_id AND l.contractor_id= _a2.contractor_id)
	 WHERE (b.dogovor_type = InDogovor Or is_all_values(InDogovor))
	 ORDER BY c.id, mc.id, xs.mid, 4
      )
      SELECT * FROM _out_table
      WHERE (month_avr + month_acc + acc_period + work_sum_period) != 0;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION report_6 IS 'Отчет по подрядчикам';

CREATE FUNCTION bldnPassport(bldnid INTEGER, bdate INTEGER, edate INTEGER) RETURNS TABLE(V01 VARCHAR(100), V02 VARCHAR(200), V03 VARCHAR(200), V04 VARCHAR(200), V05 TEXT, V06 NUMERIC(12,2), V07 DATE, V08 VARCHAR(200), V09 VARCHAR(200)) AS
$$
BEGIN
	CREATE TEMPORARY TABLE _tmp(V01 VARCHAR(100), V02 VARCHAR(200), V03 VARCHAR(200), V04 VARCHAR(200), V05 TEXT, V06 NUMERIC(12,2), V07 DATE, V08 VARCHAR(200), V09 VARCHAR(200)) ON COMMIT DROP;
	INSERT INTO _tmp
	SELECT c.name,
       	       xs.site_name || b.site_no, 
       	       CASE gwt.id WHEN 1 THEN gwt.name ELSE wt.name END,
       	       wt.name,
       	       wk.name || coalesce(' (' || w.note || ')', ''),
       	       w.work_sum,
       	       t.begin_date,
       	       w.dogovor,
       	       w.volume || ' ' || w.si
	FROM works w
	     INNER JOIN work_kinds wk ON w.workkind_id = wk.id
	     INNER JOIN work_types wt ON wk.worktype_id = wt.id
	     INNER JOIN buildings b ON w.bldn_id = b.id
	     INNER JOIN xstreets xs ON b.street_id = xs.id
	     INNER JOIN contractors c ON w.contractor_id = c.id
	     INNER JOIN terms t ON w.work_date = t.id
	     INNER JOIN global_work_types gwt on w.gwt_id = gwt.id
	     WHERE w.print_flag
	     AND t.id BETWEEN bdate AND edate
	     AND b.id = bldnid
	ORDER BY gwt.id, t.begin_date, wk.worktype_id, wk.name;
	RETURN QUERY SELECT * FROM _tmp;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION all_works(mdid INTEGER) RETURNS TABLE(V01 INTEGER, V02 TEXT, V03 VARCHAR(100), V04 TEXT, V05 INTEGER, V06 VARCHAR(50), V07 NUMERIC(12,2), V08 VARCHAR(50), V09 INTEGER) AS
$$
BEGIN
	CREATE TEMPORARY TABLE _tmp ON COMMIT DROP AS
	SELECT b.id,
       	       xs.name || ' д.' || b.bldn_no AS address,
       	       ow.work_name,
       	       ow.note,
       	       ow.work_year,
       	       ow.work_volume,
       	       ow.work_sum,
       	       ow.other_budget_note,
	       xs.mid
	FROM old_works ow
	INNER JOIN buildings b ON ow.bldn_id = b.id
	INNER JOIN xstreets xs ON b.street_id = xs.id

	UNION ALL

	SELECT b.id,
       	       xs.name || ' д.' || b.bldn_no AS address,
       	       wk.name AS work_name,
       	       w.note,
       	       CAST(EXTRACT(YEAR FROM t.begin_date) AS INTEGER) as work_year,
       	       w.volume || ' ' || w.si AS work_volume,
       	       w.work_sum,
       	       f.name AS other_budget_note,
	       xs.mid
	FROM works w
	INNER JOIN work_kinds wk ON w.workkind_id = wk.id
	INNER JOIN buildings b ON w.bldn_id = b.id
	INNER JOIN xstreets xs ON b.street_id = xs.id
	INNER JOIN terms t ON w.work_date = t.id
	INNER JOIN work_financing_sources f ON w.finance_source = f.id
	WHERE w.gwt_id > 1
	ORDER BY address, work_year, work_name;

	IF mdid <> -1002 THEN DELETE FROM _tmp WHERE mid <> mdid; END IF;

	RETURN QUERY
	SELECT * FROM _tmp;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION sub_accounts(bdate INTEGER, edate INTEGER, gwt INTEGER) RETURNS TABLE (V01 INTEGER, V02 NUMERIC(12,2), V03 DATE, V04 TEXT, V05 VARCHAR(200), V06 VARCHAR(200), V07 VARCHAR(300), V08 TEXT, volume_only VARCHAR(200), si VARCHAR(100)) AS
$$
BEGIN
	RETURN QUERY
	SELECT b.id,
	       w.work_sum,
	       t.begin_date,
	       w.volume || ' ' || w.si,
	       wk.name,
	       c.name,
	       w.dogovor,
	       w.note,
	       w.volume,
               w.si
	FROM works w
	INNER JOIN work_kinds wk ON w.workkind_id = wk.id
	INNER JOIN buildings b ON w.bldn_id = b.id
	INNER JOIN xstreets xs ON b.street_id = xs.id
	INNER JOIN contractors c ON w.contractor_id = c.id
	INNER JOIN terms t ON w.work_date = t.id
	WHERE w.gwt_id = gwt
	      AND w.finance_source = 1
	      AND w.print_flag
	      AND t.id BETWEEN bdate AND edate;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION load_subaccounts_sum(jsontext TEXT) RETURNS VOID AS
$$
  BEGIN  
    CREATE TEMPORARY TABLE _tmp ON COMMIT DROP AS
      SELECT * FROM jsonb_to_recordset(jsonText::jsonb)
		      AS t(bldn_id INTEGER, term_id INTEGER, accrued_sum NUMERIC(15,2), paid_sum NUMERIC(15,2));
    INSERT INTO sub_accounts(bldn_id, term_id, accrued_sum, paid_sum)
    SELECT * FROM _tmp
		    ON CONFLICT(bldn_id, term_id) DO
		    UPDATE SET accrued_sum = EXCLUDED.accrued_sum,         paid_sum = EXCLUDED.paid_sum;
    RETURN;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION create_sheet() RETURNS TABLE(bid INTEGER, mcname VARCHAR(200), mdname VARCHAR(200), vilname VARCHAR(200), streetname TEXT, bldnno VARCHAR(20), contname VARCHAR(200), dogname VARCHAR(300), outreport BOOLEAN) AS
$$
BEGIN
	RETURN QUERY
	SELECT b.id, mc.name, md.name, v.name, s.name || ' ' || coalesce(st.short_name, ''), b.bldn_no, c.name, d.name, b.out_report
	FROM buildings b
	     INNER JOIN streets s ON b.street_id = s.id
	     INNER JOIN street_types st ON s.street_type = st.id
	     INNER JOIN villages v ON s.village_id = v.id
	     INNER JOIN municipal_districts md ON v.md_id = md.id
	     INNER JOIN dogovors d ON b.dogovor_type = d.id
	     INNER JOIN contractors c ON b.contractor_id = c.id
	     INNER JOIN management_companies mc ON b.mc_id = mc.id
	ORDER BY mc.id, md.id, v.name, s.name, b.bldn_no;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_avr_period(accdate INTEGER, OUT termid INTEGER) AS
$$
	SELECT acc_date FROM accrueds
	WHERE acc_date = accdate;
$$ LANGUAGE SQL;

CREATE FUNCTION load_avr(bid INTEGER, contsum NUMERIC(12,2), wdate INTEGER) RETURNS VOID AS
$$
DECLARE
	cont INTEGER;
	mc INTEGER;
BEGIN
	SELECT contractor_id, mc_id INTO cont, mc FROM buildings WHERE id = bid;

	INSERT INTO accrueds(bldn_id, contractor_id, mc_id, acc_date, acc_sum) VALUES (bid, cont, mc, wdate, contsum);
	RETURN;	
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION report_7(wtid INTEGER) RETURNS TABLE(c_wkid INTEGER, c_wtname VARCHAR(200), c_wkname VARCHAR(200), c_wtid INTEGER) AS
$$
BEGIN
	CREATE TEMPORARY TABLE _tmp ON COMMIT DROP AS
	SELECT wk.id AS wk_id, wt.name AS wt_name, wk.name AS wk_name, wt.id AS wt_id
	FROM work_kinds wk
	INNER JOIN work_types wt ON wk.worktype_id = wt.id
	ORDER BY wt.id, wk.name;

	IF wtid <> -1002 THEN DELETE FROM _tmp WHERE wt_id <> wtid; END IF;

	RETURN QUERY SELECT * FROM _tmp;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION report_8(bid INTEGER, bdate INTEGER, edate INTEGER) RETURNS
TABLE (c_wtid INTEGER, c_wtname VARCHAR(200), c_note TEXT, c_address TEXT, c_site TEXT, c_bid INTEGER, c_wkname VARCHAR(200), c_year INTEGER, c_wsum NUMERIC(12,2), c_trsum NUMERIC(12,2), c_sodsum NUMERIC(12,2), c_m01 NUMERIC(12,2), c_m02 NUMERIC(12,2), c_m03 NUMERIC(12,2), c_m04 NUMERIC(12,2), c_m05 NUMERIC(12,2), c_m06 NUMERIC(12,2), c_m07 NUMERIC(12,2), c_m08 NUMERIC(12,2), c_m09 NUMERIC(12,2), c_m10 NUMERIC(12,2), c_m11 NUMERIC(12,2), c_m12 NUMERIC(12,2)) AS
$$
BEGIN
	CREATE TEMPORARY TABLE _tmp ON COMMIT DROP AS
	SELECT wt.id AS wt_id, 
	       CASE WHEN GROUPING(bldn_id)=1 THEN wt.name ELSE 'яяяяяяяяяя' END AS wt_name,
	       CASE WHEN wt.id=1 OR GROUPING(wk.name)=1 THEN '' ELSE string_agg(w.note, ',') END AS note, 
	       xs.name || ' ' || b.bldn_no AS address, 
	       xs.site_name || b.site_no AS site, 
	       w.bldn_id AS bldn_id, 
	       CASE WHEN GROUPING(wt.name)=1 THEN 'АА' ELSE wk.name END AS wk_name, 
	       CAST(EXTRACT(YEAR FROM t.begin_date) AS INTEGER) AS work_year,
	       SUM(w.work_sum) AS work_sum, 
	       SUM(CASE WHEN w.gwt_id = 2 THEN w.work_sum ELSE 0 END) AS TR,
	       SUM(CASE WHEN w.gwt_id = 1 THEN w.work_sum ELSE 0 END) AS sod,
	       SUM(CASE WHEN w.gwt_id = 1 AND EXTRACT(MONTH FROM t.begin_date) = 1 THEN w.work_sum ELSE 0 END) AS m01, 
	       SUM(CASE WHEN w.gwt_id = 1 AND EXTRACT(MONTH FROM t.begin_date) = 2 THEN w.work_sum ELSE 0 END) AS m02,  
	       SUM(CASE WHEN w.gwt_id = 1 AND EXTRACT(MONTH FROM t.begin_date) = 3 THEN w.work_sum ELSE 0 END) AS m03, 
	       SUM(CASE WHEN w.gwt_id = 1 AND EXTRACT(MONTH FROM t.begin_date) = 4 THEN w.work_sum ELSE 0 END) AS m04, 
	       SUM(CASE WHEN w.gwt_id = 1 AND EXTRACT(MONTH FROM t.begin_date) = 5 THEN w.work_sum ELSE 0 END) AS m05,  
	       SUM(CASE WHEN w.gwt_id = 1 AND EXTRACT(MONTH FROM t.begin_date) = 6 THEN w.work_sum ELSE 0 END) AS m06,
	       SUM(CASE WHEN w.gwt_id = 1 AND EXTRACT(MONTH FROM t.begin_date) = 7 THEN w.work_sum ELSE 0 END) AS m07,
	       SUM(CASE WHEN w.gwt_id = 1 AND EXTRACT(MONTH FROM t.begin_date) = 8 THEN w.work_sum ELSE 0 END) AS m08,
	       SUM(CASE WHEN w.gwt_id = 1 AND EXTRACT(MONTH FROM t.begin_date) = 9 THEN w.work_sum ELSE 0 END) AS m09,
	       SUM(CASE WHEN w.gwt_id = 1 AND EXTRACT(MONTH FROM t.begin_date) = 10 THEN w.work_sum ELSE 0 END) AS m10,
	       SUM(CASE WHEN w.gwt_id = 1 AND EXTRACT(MONTH FROM t.begin_date) = 11 THEN w.work_sum ELSE 0 END) AS m11,
	       SUM(CASE WHEN w.gwt_id = 1 AND EXTRACT(MONTH FROM t.begin_date) = 12 THEN w.work_sum ELSE 0 END) AS m12
	FROM buildings b
	     JOIN xstreets xs ON xs.id = b.street_id
	     JOIN works w ON b.id = w.bldn_id
	     JOIN work_kinds wk ON w.workkind_id = wk.id
	     JOIN work_types wt on wk.worktype_id = wt.id
	     JOIN terms t ON w.work_date = t.id
	     AND w.bldn_id = bid
	     AND w.print_flag
	     AND t.id BETWEEN bdate AND edate
	GROUP BY rollup((wt.id, wt.name), (wk.name, address, site, bldn_id, work_year));
	 RETURN QUERY SELECT * FROM _tmp order by wt_id, wt_name, wk_name;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION report_bldn_common_properties(InBldnId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS TABLE (address TEXT, site_no TEXT, bldn_cadastr VARCHAR(20), land_cadastr VARCHAR(20), builtup_year INTEGER, floors INTEGER, entrances INTEGER, has_vault BOOLEAN, flats INTEGER, live_count INTEGER, not_live_count INTEGER, structural_volume REAL, live_sq NUMERIC(12, 2), not_live_sq NUMERIC(12, 2), stairs_sq REAL, corridor_sq REAL, other_sq REAL, vault_sq REAL, attic_sq REAL, stairs_count INTEGER, has_saf BOOLEAN, has_fences BOOLEAN, bench_count INTEGER, dog_no VARCHAR(10), dog_date DATE, land_square REAL, land_builtup_square REAL, land_survey_square REAL, has_odpu BOOLEAN, has_ee BOOLEAN, has_hw BOOLEAN, has_cw BOOLEAN, has_com BOOLEAN, has_heat BOOLEAN) AS
$$
  DECLARE
    _flats_term_id INTEGER;
  BEGIN

    SELECT MAX(term_id) INTO _flats_term_id FROM flats WHERE bldn_id = InBldnId;

    RETURN QUERY
      WITH fr AS (
	SELECT
	  bldn_id
	  , SUM(square) AS f_square
	  , COUNT(flat_id)::INTEGER AS f_count
	  FROM flats
	 WHERE bldn_id = InBldnId
	   AND term_id = _flats_term_id
	   AND residental
	 GROUP BY bldn_id
      ), fnr AS (
	SELECT
	  bldn_id
	  , SUM(square) AS f_square
	  , COUNT(flat_id)::INTEGER AS f_count
	  FROM flats
	 WHERE bldn_id = InBldnId
	   AND term_id = _flats_term_id
	   AND NOT residental
	 GROUP BY bldn_id
      )
	SELECT xs.name || ' д.' || b.bldn_no,
	       xs.site_name || b.site_no,
	       b.cadastral_no,
	       bl.cadastral_no,
	       bt.built_year,
	       bt.floor_max,
	       bt.entrances,
	       (bt.vaults > 0 AND bt.vaults IS NOT NULL),
	       COALESCE(fr.f_count, 0) + COALESCE(fnr.f_count, 0),
	       fr.f_count,
	       fnr.f_count,
	       bt.structural_volume,
	       fr.f_square,
	       fnr.f_square,
	       stairs_square,
	       corridor_square,
	       other_square,
	       vaults_square,
	       attic_square,
	       stairs,
	       saf,
	       fences,
	       benches,
	       contract_no,
	       contract_date,
	       bl.use_area,
	       bl.builtup_area,
               bl.survey_area,
	       bt.has_odpu_electro OR bt.has_odpu_hotwater OR bt.has_odpu_coldwater OR bt.has_odpu_common OR bt.has_odpu_heating,
	       bt.has_odpu_electro,
	       bt.has_odpu_hotwater,
	       bt.has_odpu_coldwater,
	       bt.has_odpu_common,
	       bt.has_odpu_heating
      FROM buildings b
      INNER JOIN buildings_tech_info bt ON b.id = bt.bldn_id
      INNER JOIN buildings_land_info bl ON b.id = bl.bldn_id
      INNER JOIN xstreets xs ON b.street_id = xs.id
      LEFT JOIN fr ON fr.bldn_id = b.id
      LEFT JOIN fnr ON fnr.bldn_id = b.id
      WHERE b.id = InBldnId;
END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION report_bldn_common_properties IS 'Общее имущество дома';

CREATE FUNCTION report_9(mdid INTEGER, mcid INTEGER, contid INTEGER, dogid INTEGER, gwtid INTEGER, wtid INTEGER, wkid INTEGER, fsourceid INTEGER, bdate INTEGER, edate INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS TABLE (out_mc VARCHAR(20), out_cont VARCHAR(100), out_address TEXT, out_bid INTEGER, out_wtname VARCHAR(300), out_work TEXT, out_fsource VARCHAR(150), out_dogovor VARCHAR(150), out_volume VARCHAR, out_si VARCHAR, out_sum NUMERIC(12,2)) AS
$$
BEGIN
	CREATE TEMPORARY TABLE _tmp ON COMMIT DROP AS
	SELECT mc.report_name AS mc_name, 
		c.name AS c_name, 
		xs.name || ' д.' || b.bldn_no AS address,
		b.id AS b_id,
		wt.id AS wt_id, 
		wt.name AS wt_name, 
		wk.name || COALESCE(' (' || w.note || ')', '')  AS wk_name,
		w.dogovor AS dogovor,
		fs.name AS f_source,
		fs.id AS f_source_id,
		w.volume AS work_volume,
	        w.si AS work_si,
		w.work_sum AS work_sum,
		mc.id AS mc_id,
		md.id AS md_id,
		c.id AS cont_id,
		d.id AS dog_id,
		gwt.id AS gwt_id,
		wk.id AS wk_id
	FROM works w
	INNER JOIN buildings b ON w.bldn_id = b.id
	INNER JOIN xstreets xs ON b.street_id = xs.id
	INNER JOIN contractors c ON w.contractor_id = c.id
	INNER JOIN work_kinds wk ON w.workkind_id = wk.id
	INNER JOIN work_types wt ON wk.worktype_id = wt.id
	INNER JOIN global_work_types gwt ON w.gwt_id = gwt.id
	INNER JOIN management_companies mc ON w.mc_id = mc.id
	INNER JOIN municipal_districts md ON md.id = xs.mid
	INNER JOIN dogovors d ON b.dogovor_type = d.id
	INNER JOIN work_financing_sources fs ON w.finance_source = fs.id
	WHERE w.work_date BETWEEN bdate AND edate
	      AND (is_all_values(mdid) OR md.id = mdid)
	      AND (is_all_values(mcid) OR mc.id = mcid)
	      AND (is_all_values(contid) OR c.id = contid)
	      AND (is_all_values(dogid) OR d.id = dogid)
	      AND (is_all_values(gwtid) OR gwt.id = gwtid)
	      AND (is_all_values(wtid) OR wt.id = wtid)
	      AND (is_all_values(wkid) OR wk.id = wkid)
	      AND (is_all_values(fsourceid) OR fs.id = fsourceid);

	RETURN QUERY
	SELECT mc_name, c_name, address, b_id, wt_name, wk_name, f_source, dogovor, work_volume, work_si, work_sum
	FROM _tmp
	ORDER BY wt_name, wk_name, address;

END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION report_10(InContractorId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS TABLE(
  out_bldnid INTEGER,
  out_contname VARCHAR,
  out_mcname VARCHAR,
  out_address TEXT,
  out_totalsquare NUMERIC,
  out_cursum NUMERIC,
  out_works TEXT,
  out_worksum NUMERIC,
  out_percent NUMERIC,
  out_plansum NUMERIC,
  out_plantoyearend NUMERIC,
  out_tr NUMERIC,
  out_kr NUMERIC,
  out_ss_date DATE) AS
$$
  DECLARE
    _last_subaccount_date DATE;
    _last_subaccount_term INTEGER;
    _current_year INTEGER;
    _months INTEGER;
  BEGIN
    SELECT
      begin_date, id, EXTRACT(YEAR FROM begin_date), EXTRACT(MONTH FROM begin_date)
      INTO _last_subaccount_date, _last_subaccount_term, _current_year, _months
      FROM
	terms
     WHERE
		id = (SELECT MAX(term_id) FROM bldn_subaccounts);
    
    IF _months = 12 THEN _months = 0; _current_year = _current_year + 1; END IF;
		
    RETURN QUERY
      WITH to_year_end_works AS (
	SELECT
	  pw.bldn_id AS bid,
	  wk.name AS work_name,
	  COALESCE(NULLIF(smeta_sum, 0), work_sum) AS sums
	  FROM
	    plan_works pw
	    INNER JOIN work_kinds wk ON pw.workkind_id = wk.id
	    INNER JOIN plan_work_statuses pws ON pw.work_status = pws.id
	 WHERE pws.in_plan
	   AND EXTRACT(YEAR FROM pw.work_date)=_current_year
	   AND pw.gwt_id = 2
	 UNION ALL
	SELECT
	  w.bldn_id,
	  wk.name,
	  w.work_sum
	  FROM
	    works w
	    INNER JOIN work_kinds wk ON w.workkind_id = wk.id
	 WHERE w.work_date > (_last_subaccount_term + 1)
	   AND finance_source = 1
      ),
      year_plan_works AS (
	SELECT
	  bid AS bldn_id,
	  STRING_AGG(work_name, ', ') AS bldn_works,
	  SUM(sums) AS work_sum
	  FROM
	    to_year_end_works
	 GROUP BY bid
      ),
      expense_sum AS (
	SELECT * FROM CROSSTAB('SELECT bldn_id, expense_item, price FROM expenses WHERE expense_item in (10, 12) AND term_id = (SELECT MAX(term_id) FROM expenses) ORDER BY bldn_id, expense_item') AS ct(bldn_id INTEGER, tr NUMERIC, kr NUMERIC)
      ), t AS (
	SELECT
	  bldn_id
	  , SUM(square) AS total_square
	  FROM flats
	 WHERE term_id = (SELECT max(term_id) FROM flats)
	 GROUP by bldn_id
      )
      SELECT b.id,
      c.name,
      mc.report_name,
      xs.name || ' д. ' || b.bldn_no AS address,
      t.total_square::NUMERIC(10, 2),
      COALESCE(cs.subaccount_sum, 0),
      ypw.bldn_works,
      ypw.work_sum,
      ps.plan_percent,
      ps.plan_sum,
      (ps.plan_sum * ps.plan_percent * (12 - _months))::NUMERIC(15, 2),
      es.tr,
      es.kr,
      cs.begin_date
      FROM
      managed_buildings b
      INNER JOIN xstreets xs ON b.street_id = xs.id
      INNER JOIN management_companies mc ON mc.id = b.mc_id
      INNER JOIN t ON t.bldn_id = b.id
      INNER JOIN contractors c ON b.contractor_id = c.id
      LEFT JOIN current_subaccounts cs ON cs.bldn_id = b.id
      LEFT JOIN year_plan_works ypw ON b.id = ypw.bldn_id
      LEFT JOIN plan_subaccounts ps ON b.id = ps.bldn_id
      INNER JOIN expense_sum es ON b.id = es.bldn_id
      WHERE (c.id = InContractorId OR is_all_values(InContractorId))
      ORDER BY xs.mid, xs.vid, address;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION report_10 IS 'Планируемые субсчета';

CREATE FUNCTION add_subaccount(bid INTEGER, termId INTEGER, newsum NUMERIC(14, 2)) RETURNS VOID AS
$$
BEGIN
	INSERT INTO bldn_subaccounts (bldn_id, term_id, subaccount_sum)
	VALUES (bid, termId, newsum)
	ON CONFLICT (bldn_id, term_id) DO UPDATE
	SET subaccount_sum = EXCLUDED.subaccount_sum;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_bldn_subaccount(InBldnId INTEGER, InUserId INTEGER, InPCName VARCHAR, OUT OutSaSum NUMERIC, OUT OutSaCurSum NUMERIC, OUT OutSaDate INTEGER) AS
$$
DECLARE
	_term_id INTEGER;
	_work_sum NUMERIC;
	_addend INTEGER;
BEGIN
	SELECT MAX( term_id ) INTO _term_id FROM bldn_subaccounts WHERE bldn_id = InBldnId;
	SELECT CASE EXTRACT( month FROM begin_date ) WHEN 12 THEN 0 ELSE 1 END INTO _addend FROM terms WHERE id = _term_id;

	SELECT COALESCE( SUM( work_sum), 0) INTO _work_sum
	FROM works w
	     INNER JOIN work_financing_sources fs ON fs.id = w.finance_source
	WHERE bldn_id = InBldnId
	      AND work_date > (_term_id + _addend)
	      AND fs.from_subaccount;

	SELECT subaccount_sum, (subaccount_sum - _work_sum), term_id INTO OutSaSum, OutSaCurSum, OutSaDate
	FROM bldn_subaccounts
	WHERE bldn_id = InBldnId
	      AND term_id = _term_id;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_bldn_subaccounts(InBldnId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF bldn_subaccounts AS
$$
  SELECT * FROM bldn_subaccounts WHERE bldn_id = InBldnId
  ORDER BY term_id DESC;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_bldn_subaccounts(INTEGER, INTEGER, VARCHAR) IS 'История субсчёта дома';

CREATE FUNCTION get_bldn_subaccount_history(InBldnId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS TABLE (
  OutBldnId INTEGER,
  OutTermId INTEGER,
  OutAccruedSum NUMERIC,
  OutPaidSum NUMERIC,
  OutCurrentSum NUMERIC) AS
$$
  BEGIN
    RETURN QUERY
      SELECT sa.bldn_id,
      sa.term_id,
      sa.accrued_sum,
      sa.paid_sum,
      bs.subaccount_sum
      FROM sub_accounts AS sa
      LEFT JOIN bldn_subaccounts AS bs ON bs.bldn_id = sa.bldn_id AND bs.term_id = sa.term_id
      WHERE sa.bldn_id = InBldnId
      ORDER BY sa.term_id DESC;
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_bldn_subaccount_history IS 'История поступлений на субсчёт дома';

CREATE FUNCTION update_fact_expense(bldnId INTEGER, termId INTEGER, expenseId INTEGER, newsum NUMERIC(14, 2)) RETURNS VOID AS
$$
BEGIN
	UPDATE expenses
	SET expense_fact_sum = newsum
	WHERE bldn_id = bldnId
	      AND term_id = termId
	      AND expense_item = expenseId;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_terms(ascSort BOOLEAN, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF terms AS
$$
BEGIN
	IF ascSort THEN
	   RETURN QUERY
	   SELECT * FROM terms
	   ORDER BY begin_date ASC;
	ELSE
	   RETURN QUERY
	   SELECT * FROM terms
	   ORDER BY begin_date DESC;
	END IF;
END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_terms IS 'Получение списка периодов, отсортированного в соответствии с переданным параметром';

CREATE FUNCTION get_service_types(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF service_types AS
$$
	SELECT * FROM service_types ORDER BY id;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_service_types IS 'Список типов ЖКУ';

CREATE FUNCTION get_services(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF services AS
$$
	SELECT * FROM services ORDER BY name;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_services IS 'Список ЖКУ';

CREATE FUNCTION create_service(newname VARCHAR(100), newtype INTEGER, InPrintToPassport BOOLEAN, InUserId INTEGER, InPCName VARCHAR, OUT newid INTEGER) AS
$$
  BEGIN

    IF is_not_value(newtype)
    THEN
      INSERT INTO services(id, name, is_print_to_passport) VALUES(DEFAULT, newname, InPrintToPassport)
      RETURNING id INTO newid;
    ELSE
      INSERT INTO services(id, name, service_type, is_print_to_passport)
      VALUES(DEFAULT, newname, newtype, InPrintToPassport)
	     RETURNING id INTO newid;
    END IF;

    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 35, InUserId, InPCName, JSONB_AGG(ins_row)
      FROM services AS ins_row
     WHERE id = newid;

    PERFORM create_service_mode(newid, 'Отсутствует', InUserId, InPCName);

    INSERT INTO bldn_services(bldn_id, service_id, mode_id)
    SELECT b.id, newid, newid * 1000
      FROM buildings AS b;

  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_service IS 'Добавление ЖКУ';

CREATE FUNCTION change_service(itemid INTEGER, newname VARCHAR, newtype INTEGER, InPrintToPassport BOOLEAN, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    WITH updated_rows AS (
      UPDATE services
	 SET name = newname,
	     service_type = newtype,
	     is_print_to_passport = InPrintToPassport
       WHERE id = itemid
	     RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
	SELECT 36, InUserId, InPCName, JSONB_AGG(ttt)
	  FROM (SELECT JSONB_AGG(ss) AS prev, JSONB_AGG(updated_rows) AS upd
		  FROM services AS ss, updated_rows
		 WHERE ss.id = itemId) AS ttt;

    RETURN;
END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_service IS 'Изменение ЖКУ';

CREATE FUNCTION delete_service(itemid INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    modescount INTEGER;
  BEGIN
    SELECT COUNT(id) INTO modescount
      FROM service_modes
     WHERE service_id = itemid
       AND id != itemid * 1000;

    IF modescount > 0 THEN
	RAISE '%,%', get_error_number('has_children'), get_error_message('has_children');
    END IF;

    DELETE FROM service_modes WHERE service_id = itemid;

    WITH deleted_rows AS (
      DELETE FROM services
       WHERE id = itemid
      RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
	SELECT 37, InUserId, InPCName, JSONB_AGG(deleted_rows)
	  FROM deleted_rows;
    
    RETURN;
END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_service IS 'Удаление услуги';

CREATE FUNCTION plan_expenses_to_gis(bldnId INTEGER, bdate DATE) RETURNS TABLE (out_gisguid VARCHAR(50), out_plansum NUMERIC(12,2), out_name VARCHAR(200)) AS
$$
BEGIN
	RETURN QUERY
	SELECT ei.gis_guid, e.expense_plan_sum, ei.short_name
	FROM expense_items ei
	     INNER JOIN expenses e ON e.expense_item = ei.id
	     INNER JOIN terms t ON e.term_id = t.id
	WHERE NOT ei.gis_guid IS NULL
	      AND e.bldn_id = bldnId
	      AND bdate BETWEEN t.begin_date AND t.end_date;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION plan_price_expenses_to_gis(bldnId INTEGER, bdate DATE) RETURNS TABLE (out_gisguid VARCHAR(50), out_price NUMERIC(12,2), out_name VARCHAR(200)) AS
$$
BEGIN
	RETURN QUERY
	SELECT ei.gis_guid, e.price, ei.short_name
	FROM expense_items ei
	     INNER JOIN expenses e ON e.expense_item = ei.id
	     INNER JOIN terms t ON e.term_id = t.id
	WHERE NOT ei.gis_guid IS NULL
	      AND e.bldn_id = bldnId
	      AND bdate BETWEEN t.begin_date AND t.end_date;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_bldn_plan_subaccount(InBldnId INTEGER, InUserId INTEGER, InPcName VARCHAR) RETURNS plan_subaccounts AS
$$
	SELECT * FROM plan_subaccounts WHERE bldn_id = InBldnId;
$$ LANGUAGE SQL;

CREATE FUNCTION load_plan_subaccounts(xmlText TEXT) RETURNS VOID AS
$$
BEGIN
	UPDATE plan_subaccounts SET plan_sum = NULL;

	CREATE TEMPORARY TABLE _tmp ON COMMIT DROP AS
	SELECT XMLTABLE.* FROM XMLTABLE('//plan_subaccounts/bldn'
		       		       PASSING XMLPARSE(DOCUMENT xmlText)
				       COLUMNS bldn_id INTEGER PATH 'bldn_id',
				       plan_sum NUMERIC(14,2) PATH 'plan_subaccount');

	INSERT INTO plan_subaccounts(bldn_id, plan_sum)
	SELECT * FROM _tmp
	ON CONFLICT (bldn_id) DO UPDATE SET plan_sum = EXCLUDED.plan_sum;

	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_counter_model(itemid INTEGER) RETURNS counter_models AS
$$
	SELECT * FROM counter_models WHERE id = itemid;
$$ LANGUAGE SQL;

CREATE FUNCTION get_counter_models() RETURNS SETOF counter_models AS
$$
	SELECT * FROM counter_models ORDER BY model_name;
$$ LANGUAGE SQL;

CREATE FUNCTION create_counter_model(newname VARCHAR(300), newdti BOOLEAN, newci INTEGER, OUT newid INTEGER) AS
$$
BEGIN
	INSERT INTO counter_models(id, model_name, has_dti, calibration_interval)
	VALUES (DEFAULT, newname, newdti, newci)
	RETURNING id INTO newid;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION change_counter_model(itemid INTEGER, newname VARCHAR(300), newdti BOOLEAN, newci INTEGER) RETURNS VOID AS
$$
BEGIN
	UPDATE counter_models
	SET model_name = newname,
	    has_dti = newdti,
	    calibration_interval = newci
	WHERE id = itemid;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION delete_counter_model(itemid INTEGER) RETURNS VOID AS
$$
BEGIN
	DELETE FROM counter_models
	WHERE id = itemid;
	RETURN;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION recalc_maintenance_work(InMWorkId INTEGER) RETURNS NUMERIC AS
$$
DECLARE
	_out_work_sum NUMERIC;
BEGIN
	WITH materials_costs AS (
	     SELECT maintenance_work_id, SUM( material_cost * material_count ) AS material_sum
	     FROM works_materials
	     GROUP BY maintenance_work_id
	     HAVING maintenance_work_id = InMWorkId
	)
	UPDATE works w
	SET work_sum = ROUND( COALESCE(mc.material_sum, 0) + mw.man_hours * get_man_hour_cost_sum(w.contractor_id, mw.man_hour_mode_id, w.work_date), get_gwt_round() )
	FROM hidden_maintenance_works mw
	     LEFT JOIN materials_costs mc ON mw.id = mc.maintenance_work_id
	WHERE w.id = mw.workref_id 
	AND mw.id = InMWorkId
	RETURNING work_sum INTO _out_work_sum;
	RETURN _out_work_sum;
END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION recalc_maintenance_work IS 'Пересчет суммы работы по содержанию';

CREATE FUNCTION recalc_maintenance_works(InContractorId INTEGER, InTermId INTEGER, InCostModeId INTEGER) RETURNS VOID AS
$$
DECLARE
	_man_hour_cost NUMERIC(8, 2);
	_round INTEGER;
BEGIN
	_man_hour_cost := get_man_hour_cost_sum(InContractorId, InCostModeId, InTermId);
	_round := get_gwt_round();

	WITH tmp_works AS (
	     SELECT hmw.id AS hmw_id, w.id AS id FROM works w
	     	    INNER JOIN hidden_maintenance_works hmw ON hmw.workref_id = w.id
	     WHERE w.work_date = InTermId
	     	   AND w.contractor_id = InContractorId
		   AND hmw.man_hour_mode_id = InCostModeId
	),
	material_costs AS (
	     SELECT maintenance_work_id, SUM(material_cost * material_count) AS material_sum
	     FROM works_materials
	     WHERE maintenance_work_id IN (SELECT hmw_id FROM tmp_works)
	     GROUP BY maintenance_work_id
	)
	UPDATE works w
	SET work_sum = ROUND( COALESCE( mc.material_sum, 0 ) + mw.man_hours * _man_hour_cost, _round )
	FROM hidden_maintenance_works mw
	     LEFT JOIN material_costs mc ON mw.id = mc.maintenance_work_id
	WHERE w.id = mw.workref_id 
	AND w.id IN (SELECT id FROM tmp_works);
END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION recalc_maintenance_works IS 'Пересчет сумм работ по содержанию конкретного подрядчика в периоде с режимом';

CREATE FUNCTION get_gwt_round(OUT out_round INTEGER) AS
$$
BEGIN
	SELECT value::INTEGER INTO out_round FROM constants WHERE name = 'gwt_work_round';
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_work_materials(InMWorkId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF works_materials AS
$$
  SELECT wm.* FROM works_materials wm
  INNER JOIN work_material_types wmt ON wm.material_id = wmt.id
  WHERE maintenance_work_id = InMWorkId
  ORDER BY wmt.material_name;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_work_materials IS 'Материалы в работе по содержанию';

CREATE FUNCTION get_maintenance_work(InMWId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS maintenance_works AS
$$
  SELECT * FROM maintenance_works WHERE id = InMWId;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_maintenance_work IS 'Получение работы по содержанию по коду';

CREATE FUNCTION get_maintenance_work_by_work(InWorkId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS maintenance_works AS
$$
  SELECT * FROM maintenance_works WHERE workref_id = InWorkId;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_maintenance_work_by_work IS 'Получение работы по содержанию по коду основной работы';

CREATE FUNCTION create_maintenance_work(InWorkJSONParams TEXT, InUserId INTEGER, InPCName VARCHAR, OUT out_workid INTEGER) AS
$$
  DECLARE
    _work_info JSONB;
    _contractor INTEGER;
    _mc INTEGER;
    _man_hour_mode INTEGER;
    _material_cost NUMERIC(15, 2);
    _work_sum NUMERIC(15, 2);
    _mwork_id INTEGER;
  BEGIN
    
    IF NOT user_has_right_change(InUserId, rights_get_work_rights_number(1)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    SELECT InWorkJSONParams::JSONB->'InWorkInfo' INTO _work_info;
    
    SELECT contractor_id, mc_id INTO _contractor, _mc
      FROM buildings
     WHERE id = (_work_info->'InBldnId')::INTEGER;

    SELECT * INTO out_workid FROM create_work(1, (_work_info->'InWorkKindId')::INTEGER, (_work_info->'InDate')::INTEGER, 0, '', '', (_work_info->>'InNote')::TEXT, (_work_info->>'InPrivateNote')::TEXT, _contractor, _mc, '', 0, (_work_info->>'InPrintFlag')::BOOLEAN, (_work_info->'InBldnId')::INTEGER, InUserId, InPCName);
    
    SELECT mode_id INTO _man_hour_mode
      FROM bldn_man_hour_cost
     WHERE bldn_id = (_work_info->'InBldnId')::INTEGER
       AND term_id = (_work_info->'InDate')::INTEGER;
    
    INSERT INTO hidden_maintenance_works(man_hours, workref_id, man_hour_mode_id)
    VALUES ((_work_info->'InManHours')::NUMERIC, out_workid, _man_hour_mode)
	   RETURNING id INTO _mwork_id;
    
    WITH tmp_wmat AS (
      SELECT * FROM jsonb_populate_recordset(NULL::works_materials, InWorkJSONParams::JSONB->'InMaterials')
    )
	INSERT INTO works_materials (maintenance_work_id, material_id, material_note, material_cost, material_count, material_si)
    SELECT _mwork_id, tw.material_id, tw.material_note, tw.material_cost, tw.material_count, tw.material_si
      FROM tmp_wmat tw;

    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 27, InUserId, InPCName, JSONB_AGG(ttt)
      FROM (SELECT JSONB_AGG(hw) AS maintenance_work, JSONB_AGG(w) AS work, JSON_AGG(mat) AS materials
	      FROM hidden_maintenance_works AS hw
		   INNER JOIN works AS w ON hw.workref_id = w.id
		   INNER JOIN works_materials AS mat ON mat.maintenance_work_id = hw.id
	     WHERE hw.id = _mwork_id) AS ttt;
    
    PERFORM recalc_maintenance_work(_mwork_id);
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_maintenance_work IS 'Создание работы по содержанию';

CREATE FUNCTION change_maintenance_work(InWorkJSONParams TEXT, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    _work_info JSONB;
    _old_work_str JSONB;
    _new_work_str JSONB;
    _new_work_sum NUMERIC;
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_work_rights_number(1)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    SELECT InWorkJSONParams::JSONB->'InWorkInfo' INTO _work_info;

    SELECT JSONB_AGG(ttt) INTO _old_work_str
      FROM (SELECT JSONB_AGG(hw) AS maintenance_work, JSONB_AGG(w) AS work, JSON_AGG(mat) AS materials
	      FROM hidden_maintenance_works AS hw
		   INNER JOIN works AS w ON hw.workref_id = w.id
		   INNER JOIN works_materials AS mat ON mat.maintenance_work_id = hw.id
	     WHERE hw.id = (_work_info->'InWorkId')::INTEGER) AS ttt;

    PERFORM change_work(id, gwt_id, (_work_info->'InWorkKindId')::INTEGER, (_work_info->'InDate')::INTEGER, 0, si, volume, (_work_info->>'InNote')::TEXT, (_work_info->>'InPrivateNote')::TEXT, contractor_id, mc_id, dogovor, finance_source, (_work_info->>'InPrintFlag')::BOOLEAN, InUserId, InPCName) from works where id = (_work_info->'InWorkId')::INTEGER;

    UPDATE hidden_maintenance_works
       SET man_hours = (_work_info->'InManHours')::NUMERIC
     WHERE id = (_work_info->'InMWorkId')::INTEGER;
    
    CREATE TEMP TABLE tmp_wmat ON COMMIT DROP AS
      SELECT * FROM jsonb_populate_recordset(NULL::works_materials, InWorkJSONParams::JSONB->'InMaterials');
    
    DELETE FROM works_materials
     WHERE maintenance_work_id = (_work_info->'InMWorkId')::INTEGER
       AND NOT EXISTS (SELECT * from tmp_wmat WHERE id = works_materials.id);
    INSERT INTO works_materials (maintenance_work_id, material_id, material_note, material_cost, material_count, material_si)
    SELECT (_work_info->'InMWorkId')::NUMERIC, tw.material_id, tw.material_note, tw.material_cost, tw.material_count, tw.material_si
      FROM tmp_wmat tw
     WHERE id IS NULL;
    UPDATE works_materials 
       SET material_id = tw.material_id,
	   material_note = tw.material_note,
	   material_cost = tw.material_cost,
	   material_count = tw.material_count,
	   material_si = tw.material_si
	   FROM tmp_wmat tw
     WHERE works_materials.id = tw.id;
    
    SELECT recalc_maintenance_work((_work_info->'InMWorkId')::INTEGER) INTO _new_work_sum;

    SELECT JSONB_AGG(ttt) INTO _new_work_str
      FROM (SELECT JSONB_AGG(hw) AS maintenance_work, JSONB_AGG(w) AS work, JSON_AGG(mat) AS materials
	      FROM hidden_maintenance_works AS hw
		   INNER JOIN works AS w ON hw.workref_id = w.id
		   INNER JOIN works_materials AS mat ON mat.maintenance_work_id = hw.id
	     WHERE hw.id = (_work_info->'InMWorkId')::INTEGER) AS ttt;
    _new_work_str = JSONB_SET(_new_work_str, '{0,work_sum}', to_json(_new_work_sum)::JSONB, FALSE);

    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 28, InUserId, InPCName, JSONB_AGG(ttt)
      FROM (SELECT _old_work_str AS prev, _new_work_str AS upd) AS ttt;
    
    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_maintenance_work IS 'Изменение работы по содержанию';

CREATE FUNCTION delete_maintenance_work(InMWId INTEGER, InUserId INTEGER, InPcName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_delete(InUserId, rights_get_work_rights_number(1)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;


    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 29, InUserId, InPCName, JSONB_AGG(ttt)
      FROM (SELECT JSONB_AGG(hw) AS maintenance_work, JSONB_AGG(w) AS work, JSON_AGG(mat) AS materials
	      FROM hidden_maintenance_works AS hw
		   INNER JOIN works AS w ON hw.workref_id = w.id
		   INNER JOIN works_materials AS mat ON mat.maintenance_work_id = hw.id
	     WHERE hw.id = InMWId) AS ttt;

    DELETE FROM hidden_maintenance_works WHERE id = InMWId;
    
    DELETE FROM works_materials WHERE maintenance_work_id = InMWId;

    RETURN;
END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_maintenance_work IS 'Удаление работы по содержанию';

CREATE FUNCTION get_work_material_type(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS work_material_types AS
$$
	SELECT * FROM work_material_types WHERE id = InItemId;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION get_work_material_types(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF work_material_types AS
$$
	SELECT * FROM work_material_types ORDER BY material_name;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION create_work_material_type(InName VARCHAR, InIsTransport BOOLEAN, InUserId INTEGER, InPCName VARCHAR, OUT OutId INTEGER) AS
$$
BEGIN
	IF NOT has_access(InUserId, 'gwt', 1) THEN RAISE EXCEPTION SQLSTATE '60010'; END IF;

	INSERT INTO work_material_types (material_name, is_transport)
	VALUES (InName, InIsTransport)
	RETURNING id INTO OutId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION change_work_material_type(InItemId INTEGER, InName VARCHAR, InIsTransport BOOLEAN, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
BEGIN
	IF NOT has_access(InUserId, 'gwt', 1) THEN RAISE EXCEPTION SQLSTATE '60010'; END IF;

	IF InItemId IN (0, 1) THEN
	   RAISE EXCEPTION '99003, Нельзя изменять системный тип';
	END IF;

	UPDATE work_material_types
	SET material_name = InName,
	    is_transport = InIsTransport
	WHERE id = InItemId;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION delete_work_material_type(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
BEGIN
	IF NOT has_access(InUserId, 'gwt', 1) THEN RAISE EXCEPTION SQLSTATE '60010'; END IF;

	IF InItemId IN (0, 1) THEN
	   RAISE EXCEPTION '99003, Нельзя удалять системный тип';
	END IF;
	DELETE FROM work_material_types
	WHERE id = InItemId;
	RETURN;
EXCEPTION WHEN foreign_key_violation THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
	
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_rkc_service(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS rkc_services AS
$$
  DECLARE _out_row rkc_services%ROWTYPE;
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
    
    SELECT * INTO _out_row FROM rkc_services WHERE id = InItemId;
    RETURN _out_row;
  END;
$$ LANGUAGE plpgsql STABLE;

CREATE FUNCTION get_rkc_services(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF rkc_services AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    RETURN QUERY
    SELECT * FROM rkc_services ORDER BY name;
  END;
$$ LANGUAGE plpgsql STABLE;

CREATE FUNCTION delete_rkc_service(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_delete(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    DELETE FROM rkc_services WHERE id = InItemId;
    RETURN;

  EXCEPTION WHEN foreign_key_violation THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION create_rkc_service(InName VARCHAR, InUkServiceId INTEGER, InFullName VARCHAR, InUserId INTEGER, InPCName VARCHAR, OUT outId INTEGER) AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    INSERT INTO rkc_services(name, uk_service_id, full_name)
    VALUES (InName, InUkServiceId, InFullName)
	   RETURNING id INTO outId;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION change_rkc_service(InItemId INTEGER, InName VARCHAR, InUkServiceId INTEGER, InFullName VARCHAR, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    UPDATE rkc_services
       SET name = InName,
	   uk_service_id = InUkServiceId,
	   full_name = InFullName
     WHERE id = InItemId;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION load_rkc_values(InFileName TEXT, InTermId INTEGER, InSourceType INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    DELETE FROM rkc_values_history
     WHERE term_id = InTermId
       AND uk_accrued_source_id = InSourceType;

    EXECUTE 'COPY rkc_values_history FROM ''e:\\exchange\\postgres\\' || InFileName || ''' WITH CSV DELIMITER '';'' ENCODING ''WIN1251''';

    WITH old_occ_flats AS (
      SELECT DISTINCT rh.occ_id, rh.flat_no, rh.flat_id
	FROM rkc_values_history AS rh
	     INNER JOIN flats f ON f.flat_id = rh.flat_id
       WHERE rh.term_id = (InTermId - 1)
	     AND f.term_id = InTermId 
    )
    UPDATE rkc_values_history AS rh
	SET flat_id = of.flat_id
	FROM old_occ_flats AS of
	WHERE rh.occ_id = of.occ_id
	AND rh.flat_no = of.flat_no
	AND rh.term_id = InTermId
	AND rh.uk_accrued_source_id = InSourceType;

    UPDATE rkc_values_history AS rh
       SET flat_id = f.flat_id
	   FROM flats AS f
     WHERE rh.term_id = InTermId
       AND f.term_id = rh.term_id
       AND rh.bldn_id = f.bldn_id
       AND rh.flat_id IS NULL
       AND rh.flat_no = f.flat_no
       AND rh.uk_accrued_source_id = InSourceType;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION load_rkc_values IS 'Загрузка начисления в базу';

CREATE FUNCTION get_uk_service(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS uk_services AS
$$
  DECLARE _out_row uk_services%ROWTYPE;
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
    
    SELECT * INTO _out_row FROM uk_services WHERE id = InItemId;
    RETURN _out_row;
  END;
$$ LANGUAGE plpgsql STABLE;

CREATE FUNCTION get_uk_services(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF uk_services AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    RETURN QUERY
    SELECT * FROM uk_services ORDER BY name;
  END;
$$ LANGUAGE plpgsql STABLE;

CREATE FUNCTION delete_uk_service(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_delete(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    DELETE FROM uk_services WHERE id = InItemId;
    RETURN;

  EXCEPTION WHEN foreign_key_violation THEN
	RAISE EXCEPTION '%, Нельзя удалить таблицу, т.к. на нее есть ссылки', SQLSTATE;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION create_uk_service(InName VARCHAR, InUserId INTEGER, InPCName VARCHAR, OUT outId INTEGER) AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    INSERT INTO uk_services(name)
    VALUES (InName)
	   RETURNING id INTO outId;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION change_uk_service(InItemId INTEGER, InName VARCHAR, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    UPDATE uk_services
       SET name = InName
     WHERE id = InItemId;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_bldn_mapping(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF bldn_id_mapping AS
$$
  SELECT * FROM bldn_id_mapping;
$$ LANGUAGE SQL STABLE;


CREATE FUNCTION load_meter_readings(InFileName TEXT, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_meter_readings_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    DELETE FROM meter_readings_tmp;

    EXECUTE 'COPY meter_readings_tmp FROM ''e:\\exchange\\postgres\\' || InFileName || ''' WITH CSV DELIMITER '';'' ENCODING ''WIN1251''';

    DELETE FROM meter_readings
     WHERE term_id IN (SELECT DISTINCT(term_id) FROM meter_readings_tmp)
       AND service_id IN (SELECT DISTINCT(service_id) FROM meter_readings_tmp);

    INSERT INTO meter_readings
    SELECT bldn_id, flat_no, service_id, term_id, sum(readings)
      FROM meter_readings_tmp
     GROUP BY bldn_id, flat_no, term_id, service_id;

    DELETE FROM meter_readings_tmp;
    RETURN;
  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION get_bldn_meter_readings(InBldnId INTEGER, InServiceId INTEGER,  InUserId INTEGER, InPCName VARCHAR) RETURNS REFCURSOR AS $BF$
DECLARE
  curs1 REFCURSOR;
  _total_count INTEGER;
BEGIN
  IF NOT user_has_right_read(InUserId, rights_get_meter_readings_rights_number()) THEN
    RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
  END IF;

  SELECT COUNT(*) INTO _total_count FROM meter_readings WHERE bldn_id = InBldnId AND service_id = InServiceId;

  IF _total_count = 0 THEN
    RAISE '%, %', get_error_number('has_no_values'), get_error_message('has_no_values');
  END IF;

  OPEN curs1 FOR EXECUTE FORMAT(
    $$
    SELECT * FROM CROSSTAB (
      'SELECT flat_no , term_id , readings::NUMERIC(8,3)
      FROM meter_readings emr
      INNER JOIN buildings b on emr.bldn_id = b.id
      WHERE bldn_id = %2$s
      AND service_id = %3$s
      ORDER BY 1, 2',
      'SELECT DISTINCT term_id FROM meter_readings WHERE bldn_id = %2$s AND service_id = %3$s ORDER BY 1'
    ) AS (
      address TEXT,
      %1$s
    )
    $$,
    STRING_AGG(to_char(begin_date, 'FMmonthYYYY') || ' NUMERIC', ', '),
      InBldnId,
      InServiceId
  ) FROM (SELECT DISTINCT begin_date, id FROM terms INNER JOIN meter_readings ON term_id = id WHERE bldn_id = InBldnId AND service_id = InServiceId ORDER BY id) AS ct ;

  RETURN curs1;
END;
$BF$ LANGUAGE plpgsql;

CREATE FUNCTION get_flats_in_term_bldn(InBldnId INTEGER, InTermId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF flats AS
$$
  DECLARE
    _term_id INTEGER;
  BEGIN
    _term_id := InTermId;
    IF is_not_value(_term_id) THEN
      SELECT max(term_id) INTO _term_id FROM flats;
    END IF;

    RETURN QUERY
    SELECT * FROM flats WHERE bldn_id = InBldnId AND term_id = _term_id ORDER BY flat_no;
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_flats_in_term_bldn IS 'Список квартире в доме в определенном месяце';

CREATE FUNCTION get_flat_terms_in_bldn(InBldnId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF terms AS
$$
  SELECT DISTINCT t.* FROM terms t INNER JOIN flats f ON f.term_id = t.id WHERE f.bldn_id = InBldnId ORDER BY t.begin_date DESC;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_flat_terms_in_bldn IS 'Список месяцев в которых есть квартиры в указанном доме';

CREATE FUNCTION get_flat_history(InFlatId BIGINT, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF flats AS
$$
  SELECT * FROM flats WHERE flat_id = InFlatId ORDER BY term_id DESC;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION load_full_flats_info(InXMLText TEXT, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF flats AS
$$
  DECLARE
    myxml xml;
    _month INTEGER;
    _year INTEGER;
    _term INTEGER;
    _last_term INTEGER;
  BEGIN

    myxml := XMLPARSE(DOCUMENT InXMLText);

    SELECT (xpath('/flats/@month', myxml))[1]::TEXT::INTEGER INTO _month;
    SELECT (xpath('/flats/@year', myxml))[1]::TEXT::INTEGER INTO _year;
    SELECT id INTO _term FROM terms WHERE begin_date = make_date(_year, _month, 1);
	    
    IF _term IS NULL THEN
      RAISE exception '60001';
    END IF;

    DELETE FROM flats WHERE term_id = _term;
    SELECT MAX(term_id) INTO _last_term FROM flats;

    DROP TABLE IF EXISTS _tmp_flats;
    CREATE TEMPORARY TABLE _tmp_flats ON COMMIT DROP AS
    SELECT
      NULL::BIGINT as flat_id
      ,_term AS term_id
      ,(xpath('//bldn_id/text()', x))[1]::TEXT::INTEGER AS bldn_id
      ,(xpath('//flat_no/text()', x))[1]::TEXT::VARCHAR AS flat_no
      ,(xpath('//residental/text()', x))[1]::TEXT::INTEGER::BOOLEAN AS residental
      ,(xpath('//uninhabitable/text()', x))[1]::TEXT::INTEGER::BOOLEAN AS uninhabitable
      ,(xpath('//rooms/text()', x))[1]::TEXT::INTEGER AS rooms
      ,REPLACE((xpath('//passport_square/text()', x))[1]::TEXT, ',', '.')::NUMERIC AS passport_square
      ,REPLACE((xpath('//square/text()', x))[1]::TEXT, ',', '.')::NUMERIC AS square
      ,(xpath('//note/text()', x))[1]::TEXT AS note
      ,(xpath('//id/text()', x))[1]::TEXT::INTEGER AS tmp_flat_id
      FROM unnest(xpath('//flat', myxml)) x;

    UPDATE _tmp_flats tf
	SET flat_id = f.flat_id
	FROM flats f WHERE tf.bldn_id = f.bldn_id AND tf.flat_no = f.flat_no AND f.term_id = _last_term;

    -- Вставляем записи, у которых код квартиры найден в базе
    INSERT INTO flats(flat_id, term_id, bldn_id, flat_no, residental, uninhabitable, rooms, passport_square, square, note)
    SELECT flat_id, term_id, bldn_id, flat_no, residental, uninhabitable, rooms, passport_square, square, note FROM _tmp_flats WHERE flat_id IS NOT NULL;
    -- Вставляем записи, которых еще не было в базе
    DROP TABLE IF EXISTS _tmp_new_flats;
    CREATE TEMPORARY TABLE _tmp_new_flats ON COMMIT DROP AS
      WITH new_info AS (
	INSERT INTO flats(term_id, bldn_id, flat_no, residental, uninhabitable, rooms, passport_square, square, note)
	SELECT term_id, bldn_id, flat_no, residental, uninhabitable, rooms, passport_square, square, note FROM _tmp_flats WHERE flat_id IS NULL
																RETURNING *
      )
      SELECT * FROM new_info;

    -- Сохраняем код для новых квартир
      UPDATE _tmp_flats tf
	 SET flat_id = f.flat_id
	     FROM flats f WHERE tf.bldn_id = f.bldn_id AND tf.flat_no = f.flat_no AND f.term_id = _term AND tf.flat_id IS NULL;


    INSERT INTO flat_shares(in_period_id, flat_id, term_id, share_numerator, share_denominator, is_legal_entity, is_privatized)
    SELECT
      (xpath('//id/text()', x))[1]::TEXT::INTEGER
      ,f.flat_id
      ,_term AS term_id
      ,(xpath('//share_numerator/text()', x))[1]::TEXT::INTEGER
      ,(xpath('//share_denominator/text()', x))[1]::TEXT::INTEGER
      ,(xpath('//is_yurik/text()', x))[1]::TEXT::INTEGER::BOOLEAN
      ,(xpath('//priv/text()', x))[1]::TEXT::INTEGER::BOOLEAN
      FROM unnest(xpath('//owner', myxml)) x
	   INNER JOIN _tmp_flats f ON f.tmp_flat_id = (xpath('//flat_id/text()', x))[1]::TEXT::INTEGER;

    WITH tmo AS (
      SELECT
	(xpath('//id/text()', x))[1]::TEXT::INTEGER AS in_period_id
	,(xpath('//name/text()', x))[1]::TEXT AS name
	,(xpath('//document/text()', x))[1]::TEXT AS document
	,(xpath('//phone/text()', x))[1]::TEXT AS phone
	,(xpath('//hasPd/text()', x))[1]::TEXT::INTEGER::BOOLEAN AS has_pd
	,(xpath('//chairman/text()', x))[1]::TEXT::INTEGER::BOOLEAN AS chairman
	,(xpath('//sekretar/text()', x))[1]::TEXT::INTEGER::BOOLEAN AS sekretar
	,(xpath('//senat/text()', x))[1]::TEXT::INTEGER::BOOLEAN AS senat
	, _term AS term_id
      FROM unnest(xpath('//owner', myxml)) x
    )
    INSERT INTO owners(share_id, owner_name, owner_document, phone, has_pd_consent, is_chairman, is_sekretar, is_senat)
    SELECT
      fs.id
	, name
	, document
	, phone
	, has_pd
	, chairman
	, sekretar
	, senat
      FROM tmo
	   INNER JOIN flat_shares AS fs USING(in_period_id, term_id);

    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 106, InUserId, InPCName, JSON_BUILD_OBJECT('year', _year, 'month', _month);

    RETURN QUERY
      SELECT * FROM _tmp_new_flats;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION load_full_flats_info IS 'Загрузка xml-файла с информацией о квартирах и собственниках с удалением уже существующей информации за этот месяц';

CREATE FUNCTION get_flats_info(InBldnId INTEGER, InTermId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF flats_info AS
$$
  DECLARE
    _not_value INTEGER;
  BEGIN

    DROP TABLE IF EXISTS _tmp_saldo_sum;
    CREATE TEMP TABLE _tmp_saldo_sum ON COMMIT DROP AS
      SELECT flat_id, term_id, SUM(in_saldo-paid) AS saldo
	FROM rkc_values_history
       WHERE bldn_id = InBldnId
	 AND term_id = InTermId
       GROUP BY flat_id, term_id;

    DROP TABLE IF EXISTS _tmp_flats_info;
    CREATE TEMP TABLE _tmp_flats_info ON COMMIT DROP AS
      SELECT f.bldn_id
	     , f.term_id
	     , f.flat_id
	     , f.flat_no
	     , f.residental
	     , f.uninhabitable
	     , f.rooms
	     , f.passport_square
	     , f.square
	     , f.note
	     , f.cadastral_no
	     , fs.share_numerator
	     , fs.share_denominator
	     , fs.is_legal_entity
	     , fs.is_privatized
	     , o.id
	     , o.owner_name
	     , o.owner_document
	     , o.phone
	     , o.has_pd_consent
	     , o.is_chairman
	     , o.is_sekretar
	     , o.is_senat
	     , COALESCE(ss.saldo, 0.00) AS saldo
	     ,  CASE
	       WHEN ASCII(RIGHT(flat_no, 1)) > 57 THEN LPAD(flat_no, 10, '0')
               WHEN STRPOS(flat_no, '-') > 0 THEN LPAD(CONCAT(SUBSTR(flat_no, 1, STRPOS(flat_no, '-')-1), RIGHT(flat_no, 1)), 10, '0')
               ELSE LPAD(flat_no, 8, '0') || '00' END
	       AS sort_flat_no
	FROM flats f
	     INNER JOIN flat_shares AS fs ON fs.flat_id = f.flat_id AND f.term_id = fs.term_id
	     INNER JOIN owners AS o ON o.share_id = fs.id
	     LEFT JOIN _tmp_saldo_sum AS ss ON ss.flat_id = f.flat_id
       WHERE f.bldn_id = InBldnId
	 AND f.term_id = InTermId;
      

    IF user_has_right_read(InUserId, rights_get_owners_rights_number()) THEN
      RETURN QUERY
      SELECT * FROM _tmp_flats_info
      ORDER BY sort_flat_no, owner_name;
    ELSE
      SELECT get_not_value() INTO _not_value;
      RETURN QUERY
	SELECT DISTINCT
	bldn_id
	, term_id
	, flat_id
	, flat_no
	, residental
	, uninhabitable
	, rooms
	, passport_square
	, square
	, note
	, cadastral_no
	, _not_value
	, NULL::INTEGER
	, NULL::BOOLEAN
	, NULL::BOOLEAN
	, NULL::BIGINT
	, NULL::VARCHAR(300)
	, NULL::VARCHAR(300)
	, NULL::VARCHAR(300)
	, NULL::BOOLEAN
	, NULL::BOOLEAN
	, NULL::BOOLEAN
	, NULL::BOOLEAN
	, saldo
	, sort_flat_no
	FROM _tmp_flats_info
	ORDER BY sort_flat_no;
    END IF;
    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION get_flats_info IS 'Полная информация о помещениях в доме';

CREATE FUNCTION get_flat_history_info(InFlatId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF flats_info AS
$$
  DECLARE
    _not_value INTEGER;
  BEGIN

    DROP TABLE IF EXISTS _tmp_saldo_sum;
    CREATE TEMP TABLE _tmp_saldo_sum ON COMMIT DROP AS
      SELECT flat_id, term_id, SUM(in_saldo-paid) AS saldo
	FROM rkc_values_history
       WHERE flat_id = InFlatId
       GROUP BY flat_id, term_id;

    DROP TABLE IF EXISTS _tmp_flats_info;
    CREATE TEMP TABLE _tmp_flats_info ON COMMIT DROP AS
      SELECT f.bldn_id
	     , f.term_id
	     , f.flat_id
	     , f.flat_no
	     , f.residental
	     , f.uninhabitable
	     , f.rooms
	     , f.passport_square
	     , f.square
	     , f.note
	     , f.cadastral_no
	     , fs.share_numerator
	     , fs.share_denominator
	     , fs.is_legal_entity
	     , fs.is_privatized
	     , o.id
	     , o.owner_name
	     , o.owner_document
	     , o.phone
	     , o.has_pd_consent
	     , o.is_chairman
	     , o.is_sekretar
	     , o.is_senat
	     , COALESCE(ss.saldo, 0.00) AS saldo
	     ,  CASE
	       WHEN ASCII(RIGHT(flat_no, 1)) > 57 THEN LPAD(flat_no, 10, '0')
               WHEN STRPOS(flat_no, '-') > 0 THEN LPAD(CONCAT(SUBSTR(flat_no, 1, STRPOS(flat_no, '-')-1), RIGHT(flat_no, 1)), 10, '0')
               ELSE LPAD(flat_no, 8, '0') || '00' END
	       AS sort_flat_no
	FROM flats f
	     INNER JOIN flat_shares AS fs ON fs.flat_id = f.flat_id AND f.term_id = fs.term_id
	     INNER JOIN owners AS o ON o.share_id = fs.id
	     LEFT JOIN _tmp_saldo_sum AS ss ON ss.flat_id = f.flat_id AND ss.term_id = f.term_id
       WHERE f.flat_id = InFlatId;

    IF user_has_right_read(InUserId, rights_get_owners_rights_number()) THEN
      RETURN QUERY
      SELECT * FROM _tmp_flats_info
      ORDER BY term_id DESC, id;
    ELSE
      SELECT get_not_value() INTO _not_value;
      RETURN QUERY
	SELECT DISTINCT
	bldn_id
	, term_id
	, flat_id
	, flat_no
	, residental
	, uninhabitable
	, rooms
	, passport_square
	, square
	, note
	, cadastral_no
	, _not_value
	, NULL::INTEGER
	, NULL::BOOLEAN
	, NULL::BOOLEAN
	, NULL::BIGINT
	, NULL::VARCHAR(300)
	, NULL::VARCHAR(300)
	, NULL::VARCHAR(300)
	, NULL::BOOLEAN
	, NULL::BOOLEAN
	, NULL::BOOLEAN
	, NULL::BOOLEAN
	, saldo
	, sort_flat_no
	FROM _tmp_flats_info
	ORDER BY term_id DESC;
    END IF;
    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION get_flat_history_info IS 'Полная информация об истории помещения';

CREATE FUNCTION get_accrued_history_by_flat(InFlatId INTEGER, InUserId INTEGER, InPCName VARCHAR)
RETURNS TABLE (
  outTermId INTEGER
  , outOccId INTEGER
  , outFSourceName VARCHAR
  , outServiceName VARCHAR
  , outInSaldo NUMERIC
  , outAccrued NUMERIC
  , outAdded NUMERIC
  , outCompens NUMERIC
  , outPaid NUMERIC
  , outOutSaldo NUMERIC) AS
$$
  BEGIN
    RETURN QUERY
      SELECT rh.term_id
      , rh.occ_id
      , uas.name
      , rs.name
      , in_saldo
      , accrued
      , added
      , compens
      , paid
      , out_saldo
      FROM rkc_values_history AS rh
      INNER JOIN rkc_services AS rs ON rh.rkc_service_id = rs.id
      INNER JOIN uk_accrued_source AS uas ON rh.uk_accrued_source_id = uas.id
      WHERE flat_id = InFlatId
      ORDER BY term_id DESC, occ_id, rkc_service_id, uk_accrued_source_id;
    
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_accrued_history_by_flat IS 'История начислений по квартире';

CREATE FUNCTION get_added_types(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF rkc_added_types AS
$$
  SELECT * FROM rkc_added_types;
$$ LANGUAGE SQL STABLE;

CREATE FUNCTION load_rkc_addeds(InFileName TEXT, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    _term_id INTEGER;
    _type_id INTEGER;
    _file_version INTEGER;
    _myxml XML;

  BEGIN

    IF NOT user_has_right_change(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
    
    _myxml := XMLPARSE(DOCUMENT convert_from(pg_read_binary_file(InFileName), 'UTF8'));

    SELECT (xpath('/addeds/@term', _myxml))[1]::TEXT::INTEGER INTO _term_id;
    SELECT (xpath('/addeds/@type', _myxml))[1]::TEXT::INTEGER INTO _type_id;
    SELECT (XPATH('/addeds/@version', _myxml))[1]::TEXT::INTEGER INTO _file_version;

    IF NOT _file_version = get_file_version('rck_addeds') THEN
      RAISE '%, %', get_error_number('file_version_info'), get_error_message('file_version_info');
    END IF;

    DELETE FROM rkc_addeds_history
     WHERE term_id = _term_id AND type_id = _type_id;

    WITH rh AS (
      SELECT DISTINCT occ_id, bldn_id FROM rkc_values_history WHERE term_id = _term_id)
    INSERT INTO rkc_addeds_history
    SELECT
      _type_id
      , _term_id
      , (xpath('//occ_id/text()', x))[1]::TEXT::INTEGER
      , rh.bldn_id
      , (xpath('//service_id/text()', x))[1]::TEXT::INTEGER
      , (xpath('//sum/text()', x))[1]::TEXT::NUMERIC
      FROM UNNEST(xpath('//added', _myxml)) x
	   LEFT JOIN rh ON (xpath('//occ_id/text()', x))[1]::TEXT::INTEGER = rh.occ_id;

    RETURN;

  END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION bldn_subaccount_percent(InBldnId INTEGER, InDate DATE, InUserId INTEGER, InPCName VARCHAR, OUT OutPercent NUMERIC(4, 2)) AS
$$
  DECLARE
    _in_date DATE;
    _out_value NUMERIC(4,2);
  BEGIN

    SELECT t.begin_date INTO _in_date
      FROM terms t
	   INNER JOIN sub_accounts sa ON sa.term_id = t.id
     WHERE sa.bldn_id = InBldnId
       AND InDate BETWEEN t.begin_date AND t.end_date;
    
    IF _in_date IS NULL THEN
      SELECT max(t.begin_date) INTO _in_date
      FROM terms t
      INNER JOIN sub_accounts sa ON sa.term_id = t.id
      WHERE sa.bldn_id = InBldnId;
    END IF;
    
    SELECT GREATEST((SUM(paid_sum) / SUM(accrued_sum))::NUMERIC(4,2), 0.00) INTO OutPercent
      FROM sub_accounts
	   INNER JOIN terms t ON term_id = t.id
     WHERE t.begin_date BETWEEN (_in_date - INTERVAL '1 year') AND _in_date
       AND bldn_id = InBldnId;
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION bldn_subaccount_percent IS 'Процент поступления денег на субсчет дома';

CREATE FUNCTION get_common_property_group(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS common_property_group AS
$$
  DECLARE
    _out_row common_property_group%ROWTYPE;
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_common_property_elements_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    SELECT * INTO _out_row FROM common_property_group WHERE id = InItemId;
    RETURN _out_row;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_common_property_group IS 'получение информации из справочника групп элементов общего имущества по коду';

CREATE FUNCTION get_common_property_group_list(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF common_property_group AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_common_property_elements_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    RETURN QUERY
    SELECT * FROM common_property_group
     ORDER BY name;
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_common_property_group_list IS 'Список групп элементов общего имущества';

CREATE FUNCTION create_common_property_group(InName VARCHAR, InUserId INTEGER, InPCName VARCHAR, OUT OutId INTEGER) AS
$$
  DECLARE
    _right_id INTEGER;
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_common_property_elements_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
    
    INSERT INTO common_property_group (name)
    VALUES (InName)
	   RETURNING id INTO OutId;

    SELECT value::INTEGER + OutId INTO _right_id
      FROM constants WHERE name = 'common_property_group_prefix';
	     
    INSERT INTO access_rights(id, name)
    VALUES (_right_id, 'Элементы общего имущества: ' || InName);

    INSERT INTO roles(id, name)
    VALUES (_right_id, 'Технический отдел - ' || InName);

    UPDATE roles_access_rights
       SET access_read = TRUE
     WHERE access_id = _right_id;

    UPDATE roles_access_rights
       SET access_change = TRUE, access_delete = TRUE
	WHERE access_id = _right_id
	AND role_id = _right_id;

    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 13, InUserId, InPCName, JSONB_AGG(cpg)
      FROM common_property_group AS cpg
     WHERE id = OutId;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_common_property_group IS 'Создание группы элементов общего имущества';

CREATE FUNCTION change_common_property_group(InItemId INTEGER, InName VARCHAR, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    _right_id INTEGER;
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_common_property_elements_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    WITH t AS (
      UPDATE common_property_group
	 SET name = InName
       WHERE id = InItemId
	     RETURNING *
    )
    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 14, InUserId, InPCName,
	   JSONB_AGG(ttt)
	     FROM (SELECT JSONB_AGG(cpg) AS prev, JSONB_AGG(t) AS upd
		     FROM common_property_group AS cpg, t
		    WHERE cpg.id = InItemId) AS ttt;

    SELECT value::INTEGER + InItemId INTO _right_id
      FROM constants WHERE name = 'common_property_group_prefix';

    UPDATE access_rights
       SET name = 'Элементы общего имущества: ' || InName
     WHERE id = _right_id;

    UPDATE roles
       SET name = 'Технический отдел - ' || InName
     WHERE id = _right_id;
			  
    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_common_property_group IS 'Изменение группы элементов общего имущества';

CREATE FUNCTION delete_common_property_group(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    _right_id INTEGER;
  BEGIN
    IF NOT user_has_right_delete(InUserId, rights_get_common_property_elements_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    WITH deleted_rows AS (
      DELETE FROM common_property_group
       WHERE id = InItemId
      RETURNING *
    )
    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 15, InUserId, InPCName, JSONB_AGG(deleted_rows)
      FROM deleted_rows;

    SELECT value::INTEGER + InItemId INTO _right_id
      FROM constants WHERE name = 'common_property_group_prefix';

    DELETE FROM access_rights
     WHERE id = _right_id;
    DELETE FROM roles
     WHERE id = _right_id;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_common_property_group IS 'Удаление группы элементов общего имущества';

CREATE FUNCTION get_common_property_element(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS common_property_element AS
$$
  DECLARE
    _out_row common_property_element%ROWTYPE;
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_common_property_elements_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    SELECT * INTO _out_row FROM common_property_element WHERE id = InItemId;
    RETURN _out_row;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_common_property_element IS 'получение информации из справочника элементов общего имущества по коду';

CREATE FUNCTION get_common_property_element_list(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF common_property_element AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_common_property_elements_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    RETURN QUERY
    SELECT * FROM common_property_element
     ORDER BY group_id, name;
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_common_property_element_list IS 'Список элементов общего имущества';

CREATE FUNCTION create_common_property_element(InGroupId INTEGER, InName VARCHAR, InRequired BOOLEAN, InUserId INTEGER, InPCName VARCHAR, OUT OutId INTEGER) AS
$$
  BEGIN

    IF NOT user_has_right_change(InUserId, rights_get_common_property_group_rights_number(InGroupId)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    INSERT INTO common_property_element (group_id, name, is_required)
    VALUES (InGroupId, InName, InRequired)
	   RETURNING id INTO OutId;
    
    INSERT INTO building_common_property_elements(element_id, bldn_id)
    SELECT OutId, id
      FROM buildings;

    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 19, InUserId, InPCName, JSONB_AGG(cpe)
      FROM common_property_element AS cpe
     WHERE id = OutId;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_common_property_element IS 'Создание элемента общего имущества';

CREATE FUNCTION change_common_property_element(InItemId INTEGER, InGroupId INTEGER, InName VARCHAR, InRequired BOOLEAN, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN

    IF NOT user_has_right_change(InUserId, rights_get_common_property_group_rights_number(InGroupId)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    WITH updated AS (
      UPDATE common_property_element
	 SET name = InName,
	     group_id = InGroupId,
	     is_required = InRequired
       WHERE id = InItemId
	     RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 20, InUserId, InPCName,
	   JSONB_AGG(ttt)
      FROM (SELECT JSONB_AGG(cpe) AS prev, JSONB_AGG(updated) AS upd
	      FROM common_property_element AS cpe, updated
	     WHERE cpe.id = InItemId) AS ttt;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_common_property_element IS 'Изменение элемента общего имущества';

CREATE FUNCTION delete_common_property_element(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    _group_id INTEGER;

  BEGIN

    SELECT group_id INTO _group_id FROM common_property_element WHERE id = InItemId;
    IF NOT user_has_right_delete(InUserId, rights_get_common_property_group_rights_number(_group_id)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    IF EXISTS (SELECT DISTINCT(bldn_id) FROM building_common_property_elements_history WHERE element_id = InItemId AND is_contain) THEN
      RAISE '%, %', get_error_number('delete_denied'), get_error_message('delete_denied');
    END IF;
    
    WITH deleted_rows AS (
      DELETE FROM common_property_element
       WHERE id = InItemId
      RETURNING *
    )
    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 21, InUserId, InPCName, JSONB_AGG(deleted_rows)
      FROM deleted_rows;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_common_property_element IS 'Удаление элемента общего имущества';

CREATE FUNCTION change_bldn_common_property_element_state(InBldnId INTEGER, InElementId INTEGER, InNewState BOOLEAN, InUserId INTEGER, InPCName VARCHAR) RETURNS void AS
$$
  DECLARE
    _current_state BOOLEAN;
    _element_group INTEGER;
    _parameters JSONB;

  BEGIN

    SELECT is_contain INTO _current_state
      FROM building_common_property_elements
     WHERE bldn_id = InBldnId
       AND element_id = InElementId;

    IF _current_state = InNewState THEN RETURN; END IF;

    SELECT group_id INTO _element_group
      FROM common_property_element
     WHERE id = InElementId;

    IF NOT user_has_right_change(InUserId, rights_get_common_property_group_rights_number(_element_group)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    UPDATE building_common_property_elements
       SET is_contain = InNewState
     WHERE bldn_id = InBldnId
       AND element_id = InElementId;

    IF InNewState THEN
      INSERT INTO building_common_property_element_parameter
      SELECT InBldnId, cpep.id
      FROM common_property_element_parameter AS cpep
      WHERE cpep.element_id = InElementId;
    ELSE
      UPDATE building_common_property_elements
	 SET element_state = ''
       WHERE bldn_id = InBldnId
	 AND element_id = InElementId;

      WITH deleted AS (
	DELETE FROM building_common_property_element_parameter
	 WHERE bldn_id = InBldnId
	   AND parameter_id IN (SELECT id FROM common_property_element_parameter WHERE element_id = InElementId)
	RETURNING *
      ) SELECT JSONB_AGG(deleted) INTO _parameters FROM deleted;
    END IF;

    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 22, InUserId, InPCName, JSONB_AGG(ttt)
      FROM (SELECT JSONB_AGG(bce) AS elem, _parameters AS params
	      FROM building_common_property_elements AS bce
	     WHERE bldn_id = InBldnId
	       AND element_id = InElementId) AS ttt;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_bldn_common_property_element_state IS 'Изменение состояния элемента общего имущества в доме';

CREATE FUNCTION change_bldn_common_property_element_value(InBldnId INTEGER, InElementId INTEGER, InValue VARCHAR, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    _element_group INTEGER;

  BEGIN
    IF NOT (SELECT is_contain FROM building_common_property_elements WHERE bldn_id = InBldnId AND element_id = InElementId) THEN 
      RETURN;
    END IF;

    SELECT group_id INTO _element_group FROM common_property_element WHERE id = InElementId;
    IF NOT user_has_right_change(InUserId, rights_get_common_property_group_rights_number(_element_group)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    WITH updated AS (
      UPDATE building_common_property_elements
	 SET element_state = InValue
       WHERE bldn_id = InBldnId
	 AND element_id = InElementId
	     RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
	SELECT 22, InUserId, InPCName, JSONB_AGG(ttt)
	  FROM (SELECT JSONB_AGG(be) AS prev, JSONB_AGG(updated) AS upd
		  FROM building_common_property_elements AS be
		       INNER JOIN updated ON  be.bldn_id = updated.bldn_id AND be.element_id = updated.element_id) AS ttt;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_bldn_common_property_element_value IS 'Изменение технического состояния элемента общего имущества в доме';

CREATE FUNCTION change_bldn_common_property_parameter_value(InBldnId INTEGER, InParameterId BIGINT, InValue TEXT, InUserId INTEGER, InPCName VARCHAR) RETURNS void AS
$$
  DECLARE
    _element_group INTEGER;

  BEGIN
    IF NOT (SELECT is_using FROM common_property_element_parameter WHERE id = InParameterId) THEN
      RETURN;
    END IF;

    SELECT group_id INTO _element_group FROM common_property_dictionary WHERE parameter_id = InParameterId;
    IF NOT user_has_right_change(InUserId, rights_get_common_property_group_rights_number(_element_group)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    WITH updated AS (
      UPDATE building_common_property_element_parameter
	 SET parameter_value = InValue
       WHERE bldn_id = InBldnId
	 AND parameter_id = InParameterId
	     RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
	SELECT 23, InUserId, InPCName, JSONB_AGG(ttt)
	  FROM (SELECT JSONB_AGG(bp) AS prev, JSONB_AGG(updated) AS upd
		  FROM building_common_property_element_parameter AS bp
		       INNER JOIN updated ON  bp.bldn_id = updated.bldn_id AND bp.parameter_id = updated.parameter_id) AS ttt;

    RETURN;

END;    
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_bldn_common_property_parameter_value IS 'Изменение значения параметра общего имущества в доме';

CREATE FUNCTION get_bldn_common_properties(InBldnId INTEGER, ShowNotRequired BOOLEAN, InUserId INTEGER, InPCName VARCHAR)
  RETURNS TABLE(
    OutRank TEXT
    , OutGroupId INTEGER
    , OutElementId INTEGER
    , OutParameterId BIGINT
    , OutName VARCHAR
    , OutState TEXT
    , OutUsing BOOLEAN
  ) AS
$$
  BEGIN
    RETURN QUERY
    WITH l1 AS (
      SELECT CHAR_LENGTH(MAX(id)::TEXT) AS lg FROM common_property_group
    ), l2 AS (
      SELECT CHAR_LENGTH(MAX(id)::TEXT) AS le FROM common_property_element
    ), t1 AS (
      SELECT LPAD((RANK() OVER (ORDER BY name))::TEXT, lg, '0') AS sort, id AS group, 0 AS element, 0 AS parameter, name, '' AS state_text, true AS using_param FROM common_property_group, l1
    ), t2 AS (
      SELECT sort || '.' || LPAD((RANK() OVER (PARTITION BY ce.group_id ORDER BY ce.name))::TEXT, le, '0') AS s2, 0, ce.id, 0, ce.name , CASE WHEN is_contain THEN element_state ELSE  'нет' END, is_contain FROM building_common_property_elements AS be INNER JOIN  common_property_element AS ce ON be.element_id = ce.id INNER JOIN t1 ON ce.group_id = t1.group, l2 WHERE be.bldn_id = InBldnId AND (ShowNotRequired OR ce.is_required OR is_contain)
    ), t3 AS (
      SELECT s2 || '.' || RANK() OVER (PARTITION BY element_id ORDER BY cp.name), 0, 0, cp.id, cp.name, parameter_value, is_using FROM building_common_property_element_parameter AS bp INNER JOIN common_property_element_parameter AS cp ON bp.parameter_id = cp.id INNER JOIN t2 ON cp.element_id = t2.id WHERE bp.bldn_id = InBldnId
    )
      SELECT * FROM t1
      UNION ALL
      SELECT * FROM t2
      UNION ALL
      SELECT * FROM t3
      ORDER BY 1;
    
    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION get_bldn_common_properties IS 'Элементы общего имущества в доме';

CREATE FUNCTION load_offers_works(InFileName VARCHAR, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    _myxml xml;
    _offer_date DATE;
    _file_version INTEGER;
  BEGIN

    IF NOT user_has_right_change(InUserId, rights_get_offers_work_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    _myxml := XMLPARSE(DOCUMENT CONVERT_FROM(PG_READ_BINARY_FILE(InFileName), 'UTF8'));

    SELECT (XPATH('/buildings/@version', _myxml))[1]::TEXT::INTEGER INTO _file_version;

    IF NOT _file_version = get_file_version('offers') THEN
      RAISE '%, %', get_error_number('file_version_info'), get_error_message('file_version_info');
    END IF;

    SELECT make_date((xpath('/buildings/@year', _myxml))[1]::TEXT::INTEGER, 1, 1) INTO _offer_date;

    WITH t AS (
      DELETE FROM offers_work WHERE offers_year = _offer_date
      RETURNING *
    )
	INSERT INTO log_offers (bldn_id, action_id, user_id, pc_name, action_description)
    SELECT
      DISTINCT t.bldn_id
      , 3
      , InUserId
      , InPCName
      , 'Очистка данных'
      FROM t;

    WITH t AS (
      INSERT INTO offers_work (bldn_id, offers_year, work_name, work_sum, priority)
      SELECT
	(xpath('//bldnid//text()', x))[1]::TEXT::INTEGER
	, _offer_date
	, UNNEST(xpath('//work/name/text()', x))::TEXT
	, UNNEST(xpath('//work/sum/text()', x))::TEXT::NUMERIC
	, UNNEST(xpath('//work/priority/text()', x))::TEXT::INTEGER
	FROM UNNEST(xpath('//bldn', _myxml)) x
	       RETURNING *
    )
	INSERT INTO log_offers(bldn_id, action_id, user_id, pc_name, action_description)
	SELECT
	  t.bldn_id
	  , 1
	  , InUserId
	  , InPCName
	  , CONCAT_WS(' :-: ', 'Загрузка из файла', t.work_name, t.work_sum, t.priority, t.offers_year)
	  FROM t;

  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION load_offers_works IS 'Загрузка предложений из файла';

CREATE FUNCTION load_offers_expenses(InFileName VARCHAR, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    myxml xml;
    _offer_date DATE;
  BEGIN

    IF NOT user_has_right_change(InUserId, rights_get_offers_work_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    myxml := XMLPARSE(DOCUMENT CONVERT_FROM(PG_READ_BINARY_FILE(InFileName), 'UTF8'));

    SELECT make_date((xpath('/buildings/@year', myxml))[1]::TEXT::INTEGER, 1, 1) INTO _offer_date;

    DELETE FROM offers_expense WHERE offers_date = _offer_date;

    INSERT INTO offers_expense (bldn_id, offers_date, expense_item, expense_value)
    SELECT
      (xpath('//bldnid//text()', x))[1]::TEXT::INTEGER
      , _offer_date
      , UNNEST(xpath('//expense/expenseid/text()', x))::TEXT::INTEGER
      , UNNEST(xpath('//expense/new/text()', x))::TEXT::NUMERIC
      FROM UNNEST(xpath('//bldn', myxml)) x;

  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION load_offers_expenses IS 'Загрузка предложений по структуре из файла';

CREATE FUNCTION get_offers_work(InItemId BIGINT, InUserId INTEGER, InPCName VARCHAR) RETURNS offers_work AS
$$
  DECLARE
    _out_row offers_work%ROWTYPE;
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_offers_work_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    SELECT * INTO _out_row FROM offers_work WHERE id = InItemId;

    RETURN _out_row;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_offers_work IS 'Предлагаемая работа по коду';

CREATE FUNCTION create_offers_work(InBldnId INTEGER, InWorkName VARCHAR, InWorkSum NUMERIC, InWorkPriority INTEGER, InOfferYear DATE, InUserId INTEGER, InPCName VARCHAR, OUT OutId BIGINT) AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_offers_work_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
    
    INSERT INTO offers_work(bldn_id, work_name, work_sum, priority, offers_year)
    VALUES (InBldnId, InWorkName, InWorkSum, InWorkPriority, InOfferYear)
	   RETURNING id INTO OutId;
    
    INSERT INTO log_offers(bldn_id, action_id, user_id, pc_name, action_description)
    VALUES (
      InBldnId
      , 1
      , InUserId
      , InPCName
      , CONCAT_WS(' :-: ', InWorkName, InWorkSum, InWorkPriority, InOfferYear)
    );

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_offers_work IS 'Создание предлагаемой работы';


CREATE FUNCTION change_offers_work(InItemId BIGINT, InWorkName VARCHAR, InWorkSum NUMERIC, InWorkPriority INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_offers_work_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    WITH t AS (
      UPDATE offers_work
	 SET work_name = InWorkName
	     , work_sum = InWorkSum
	     , priority = InWorkPriority
       WHERE id = InItemId
	     RETURNING *
    )
	INSERT INTO log_offers(bldn_id, action_id, user_id, pc_name, action_description)
    SELECT
      bldn_id
      , 2
      , InUserId
      , InPCName
      , CONCAT_WS(' :-: ', 'Было', work_name, work_sum, priority, 'Стало', InWorkName, InWorkSum, InWorkPriority)
      FROM offers_work WHERE id = InItemId;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_offers_work IS 'Изменение предлагаемой работы';

CREATE FUNCTION delete_offers_work(InItemId BIGINT, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_delete(InUserId, rights_get_offers_work_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    WITH deleted_rows AS (
      DELETE FROM offers_work WHERE id = InItemId
      RETURNING *
    )
	INSERT INTO log_offers(bldn_id, action_id, user_id, pc_name, action_description)
    SELECT
      bldn_id
      , 3
      , InUserId
      , InPCName
      , CONCAT_WS(' :-: ', work_name, work_sum, priority)
      FROM deleted_rows;

    RETURN;
	
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_offers_work IS 'Удаление предлагаемой работы';

CREATE FUNCTION get_work_offers_in_bldn(InBldnId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF offers_work AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_offers_work_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    RETURN QUERY
    SELECT * FROM offers_work WHERE bldn_id = InBldnId ORDER BY priority, work_name;
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_work_offers_in_bldn IS 'Предлагаемые работы по дому';

CREATE FUNCTION get_work_annex_in_bldn(InBldnId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF offers_annex AS
$$
  SELECT * FROM offers_annex WHERE bldn_id = InBldnId ORDER BY annex_date DESC;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_work_annex_in_bldn IS 'История предложений по дому';

CREATE FUNCTION get_work_annex(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS XML AS
$$
  SELECT offer_text FROM offers_annex WHERE id = InItemId;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_work_annex IS 'Вывод предложения из истории';

CREATE FUNCTION create_and_get_work_annex(InBldnId INTEGER, InUserId INTEGER, InPCName VARCHAR, OUT OutXml XML) AS
$$
  DECLARE
    _out_xml XML;
    _has_doorcloser BOOLEAN;
    _has_doorphone BOOLEAN;
    _has_odpu BOOLEAN;
    _dogovor INTEGER;
    _has_thermoregulator BOOLEAN;
  BEGIN

    SELECT has_doorphone
	   , has_doorcloser
	   , has_thermoregulator
	   , has_odpu_common OR has_odpu_heating OR has_odpu_hotwater
      INTO _has_doorphone, _has_doorcloser, _has_thermoregulator, _has_odpu
      FROM buildings_tech_info WHERE bldn_id = InBldnId;

    SELECT dogovor_type INTO _dogovor
      FROM buildings WHERE id = InBldnId;

    DROP TABLE IF EXISTS _tmp_annex;
    CREATE TEMP TABLE _tmp_annex ON COMMIT DROP AS
      SELECT * FROM offers_work WHERE bldn_id = InBldnId;
    IF _has_doorphone THEN
      INSERT INTO _tmp_annex(bldn_id, work_name, priority)
      VALUES (InBldnId, 'Ремонт домофонов (по необходимости)', 30);
    END IF;
    IF _has_doorcloser THEN
      INSERT INTO _tmp_annex(bldn_id, work_name, priority)
      VALUES (InBldnId, 'Замена доводчиков на дверях (по необходимости)', 40);
    END IF;
   IF _has_odpu THEN
      INSERT INTO _tmp_annex(bldn_id, work_name, priority)
      VALUES (InBldnId, 'Поверка, замена, ремонт комплектующих ОДПУ ТЭ (по необходимости)', 50);
   END IF;
    IF _has_thermoregulator THEN
      INSERT INTO _tmp_annex(bldn_id, work_name, priority)
      VALUES (InBldnId, 'Ремонт узла автоматического регулирования', 60);
    END IF;
    IF _dogovor = 2 THEN
      INSERT INTO _tmp_annex(bldn_id, work_name, priority)
      VALUES (InBldnId, 'Прочие услуги по благоустройству (по необходимости)', 70);
    END IF;
 
    SELECT QUERY_TO_XML('SELECT * FROM _tmp_annex ORDER BY priority', FALSE, FALSE, '') INTO outXml;

    INSERT INTO offers_annex(bldn_id, offer_text, user_id, pc_name)
    SELECT InBldnId
	   , _out_xml
	   , InUserId
	   , InPCName
	     ;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_and_get_work_annex IS 'Вывести предложения по дому и сохранить его';

CREATE FUNCTION get_expense_group(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS expense_groups AS
$$
  DECLARE
    _out_row expense_groups%ROWTYPE;
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    SELECT * INTO _out_row FROM expense_groups WHERE id = InItemId;

    RETURN _out_row;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_expense_group IS 'Группа структуры платы по коду';

CREATE FUNCTION create_expense_group(InName VARCHAR, InReportPriority INTEGER, InParentGroup INTEGER, InUserId INTEGER, InPCName VARCHAR, OUT OutId INTEGER) AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
    
    INSERT INTO expense_groups(name, report_priority, parent_group)
    VALUES (InName, InReportPriority, InParentGroup)
	   RETURNING id INTO OutId;

    INSERT INTO log_log (action_id, user_id, pc_name, action_description)
    VALUES (
      7,
      InUserId,
      InPCName,
      CONCAT_WS(' :-: ', InName, InReportPriority, InParentGroup)
    );

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_expense_group IS 'Создание группы структуры';

CREATE FUNCTION change_expense_group(InItemId INTEGER, InName VARCHAR, InReportPriority INTEGER, InParentGroup INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    WITH t AS (
      UPDATE expense_groups
	 SET name = InName
	     , report_priority = InReportPriority
	     , parent_group = InParentGroup
       WHERE id = InItemId
	     RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, action_description)
    SELECT
      8
      , InUserId
      , InPCName
      , CONCAT_WS(' :-: ', 'Было', name, report_priority, parent_group, 'Стало', InName, InReportPriority, InParentGroup)
      FROM expense_groups WHERE id = InItemId;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_expense_group IS 'Изменение группы структуры платы';

CREATE FUNCTION delete_expense_group(InItemId BIGINT, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_delete(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    WITH deleted_rows AS (
      DELETE FROM expense_groups WHERE id = InItemId
      RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, action_description)
    SELECT
      9
      , InUserId
      , InPCName
      , CONCAT_WS(' :-: ', name, report_priority, parent_group)
      FROM deleted_rows;

    RETURN;
	
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_expense_group IS 'Удаление группы структуры платы';

CREATE FUNCTION get_expense_groups(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF expense_groups AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_bldn_accrued_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    RETURN QUERY
      WITH t AS (
	SELECT id, report_priority
	  FROM expense_groups)
    SELECT eg.*
      FROM expense_groups eg LEFT JOIN t ON eg.parent_group = t.id
      ORDER BY CASE COALESCE(eg.parent_group, 0) WHEN 0 THEN eg.report_priority * 1000
	       ELSE t.report_priority * 1000 + eg.report_priority
		 END,
    name;
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_expense_groups IS 'Список групп структуры платы';

CREATE FUNCTION get_common_property_element_parameter(InItemId BIGINT, InUserId INTEGER, InPCName VARCHAR) RETURNS common_property_element_parameter AS
$$
  DECLARE
    _out_row common_property_element_parameter%ROWTYPE;
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_common_property_elements_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    SELECT * INTO _out_row FROM common_property_element_parameter WHERE id = InItemId;
    RETURN _out_row;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_common_property_element_parameter IS 'получение информации из справочника параметров элементов общего имущества по коду';

CREATE FUNCTION get_common_property_element_parameters_list(InElementId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF common_property_element_parameter AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_common_property_elements_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    RETURN QUERY
      SELECT cpp.*
      FROM common_property_element_parameter AS cpp
      INNER JOIN common_property_element AS cpe ON cpp.element_id = cpe.id
      WHERE (element_id = InElementId OR is_all_values(InElementId))
      ORDER BY cpe.group_id, cpe.name, cpp.name;
    
    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION get_common_property_element_parameters_list IS 'Список параметров элемента общего имущества';

CREATE FUNCTION create_common_property_element_parameter(InElementId INTEGER, InName VARCHAR, InIsUsing BOOLEAN, InUserId INTEGER, InPCName VARCHAR, OUT OutId BIGINT) AS
$$
  DECLARE
    _elem_group_id INTEGER;

  BEGIN

    SELECT group_id INTO _elem_group_id FROM common_property_element WHERE id = InElementId;
    IF NOT user_has_right_change(InUserId, rights_get_common_property_group_rights_number(_elem_group_id)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    INSERT INTO common_property_element_parameter(element_id, name, is_using)
    VALUES (InElementId, InName, InIsUsing)
	   RETURNING id INTO OutId;

    INSERT INTO building_common_property_element_parameter(parameter_id, bldn_id)
    SELECT OutId, bldn_id
      FROM building_common_property_elements
	     WHERE element_id = InElementId
	       AND is_contain;

    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 16, InUserId, InPCName, JSONB_AGG(common_property_dictionary)
      FROM common_property_dictionary
     WHERE parameter_id = OutId;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_common_property_element_parameter IS 'Создание параметра элемента общего имущества дома';

CREATE FUNCTION change_common_property_element_parameter(InItemId INTEGER, InName VARCHAR, InIsUsing BOOLEAN, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    _elem_group_id INTEGER;
  BEGIN

    SELECT group_id INTO _elem_group_id
      FROM common_property_element AS cpe
	   INNER JOIN common_property_element_parameter AS cpar ON cpar.element_id = cpe.id
     WHERE cpar.id = InItemId;
    IF NOT user_has_right_change(InUserId, rights_get_common_property_group_rights_number(_elem_group_id)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    WITH updated AS (
    UPDATE common_property_element_parameter
       SET name = InName,
	   is_using = InIsUsing
     WHERE id = InItemId
	   RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 17, InUserId, InPCName, JSONB_AGG(ttt)
      FROM (SELECT *, InName AS new_name, InIsUsing AS new_is_using
	      FROM common_property_dictionary
	     WHERE parameter_id = InItemId) AS ttt;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_common_property_element_parameter IS 'Изменение параметра элемента общего имущества дома';

CREATE FUNCTION delete_common_property_element_parameter(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    _elem_group_id INTEGER;

  BEGIN

    -- проверка прав пользователя
    SELECT group_id INTO _elem_group_id
      FROM common_property_element AS cpe
	   INNER JOIN common_property_element_parameter AS cpar ON cpar.element_id = cpe.id
     WHERE cpar.id = InItemId;
    IF NOT user_has_right_delete(InUserId, rights_get_common_property_group_rights_number(_elem_group_id)) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    -- если есть история, то удаление запрещено
    IF EXISTS (SELECT DISTINCT(bldn_id) FROM building_common_property_element_parameter_history WHERE parameter_id = InItemId) THEN
      RAISE '%, %', get_error_number('delete_denied'), get_error_message('delete_denied');
    END IF;

    WITH deleted AS (
    DELETE FROM common_property_element_parameter
     WHERE id = InItemId
    RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 18, InUserId, InPCName, JSONB_AGG(ttt)
      FROM (SELECT cpg.name AS group_name
		   , cpg.id AS group_id
		   , cpe.name AS element_name
		   , cpe.id AS element_id
		   , d.id AS parameter_id
		   , d.name AS parameter_name
	      FROM deleted AS d
		   INNER JOIN common_property_element AS cpe ON d.element_id = cpe.id
		   INNER JOIN common_property_group AS cpg ON cpe.group_id = cpg.id) AS ttt;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_common_property_element_parameter IS 'Удаление параметра элемента общего имущества дома';

CREATE FUNCTION create_man_hour_cost_mode(InName VARCHAR, InUserId INTEGER, InPCName VARCHAR, OUT OutId INTEGER) AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_contractor_cost_rights_number()) AND InUserId != 0 THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;
    
    INSERT INTO man_hour_cost_modes(name)
    VALUES (InName)
	   RETURNING id INTO OutId;

    INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 24, InUserId, InPCName, JSONB_AGG(man_hour_cost_modes)
      FROM man_hour_cost_modes
     WHERE id = OutId;

    INSERT INTO man_hour_cost_rates(mode_id, term_id, contractor_id)
    SELECT OutId, terms.id, cont.id
      FROM terms,
	   contractors AS cont
     WHERE terms.id = (SELECT id FROM terms WHERE begin_date = (SELECT MAX(begin_date) FROM terms))
       AND (cont.bldn_contractor AND cont.is_using);

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION create_man_hour_cost_mode IS 'Создание режима стоимости человекочаса';

CREATE FUNCTION change_man_hour_cost_mode(InItemId INTEGER, InName VARCHAR, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_contractor_cost_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    IF InItemId = 0 THEN
      RAISE '%, %', get_error_number('delete_denied'), get_error_message('delete_denied');
    END IF;

    WITH t AS (
      UPDATE man_hour_cost_modes
	 SET name = InName
       WHERE id = InItemId
	     RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 25, InUserId, InPCName,
	   JSONB_AGG(ttt)
      FROM (SELECT JSONB_AGG(st) AS prev, JSONB_AGG(t) AS upd
	      FROM man_hour_cost_modes AS st, t
	     WHERE st.id = InItemId) AS ttt;
    
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_man_hour_cost_mode IS 'Изменение режима стоимости человекочаса';

CREATE FUNCTION delete_man_hour_cost_mode(InItemId BIGINT, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_contractor_cost_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    IF InItemId = 0 THEN
      RAISE '%, %', get_error_number('delete_denied'), get_error_message('delete_denied');
    END IF;

    IF EXISTS (SELECT mode_id FROM bldn_man_hour_cost WHERE mode_id = InItemId AND term_id = (SELECT MAX(term_id) FROM bldn_man_hour_cost)) THEN
      RAISE '%, %', get_error_number('has_children'), get_error_message('has_children');
    END IF;

    WITH deleted_rows AS (
      DELETE FROM man_hour_cost_modes WHERE id = InItemId
      RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
    SELECT 26, InUserId, InPCName, JSONB_AGG(deleted_rows)
      FROM deleted_rows;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION delete_man_hour_cost_mode IS 'Удаление режима стоимости человекочаса';

CREATE FUNCTION get_man_hour_cost_mode(InItemId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS man_hour_cost_modes AS
$$
  DECLARE
    _out_row man_hour_cost_modes%ROWTYPE;
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_contractor_cost_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    SELECT * INTO _out_row FROM man_hour_cost_modes WHERE id = InItemId;

    RETURN _out_row;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_man_hour_cost_mode IS 'Режим стоимости человекочаса по коду';

CREATE FUNCTION get_man_hour_cost_modes(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF  man_hour_cost_modes AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_contractor_cost_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    RETURN QUERY
      SELECT * FROM man_hour_cost_modes
      ORDER BY name;

    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_man_hour_cost_modes IS 'Список режимом стоимости человекочаса';

CREATE FUNCTION get_man_hour_cost_sum(InContractorId INTEGER, InModeId INTEGER, InTermId INTEGER, OUT out_cost_sum NUMERIC(8, 2)) AS
$$
  BEGIN
    SELECT cost_sum INTO out_cost_sum
      FROM man_hour_cost_rates
     WHERE mode_id = InModeId
       AND contractor_id = InContractorId
       AND term_id = InTermId;
    
    IF out_cost_sum IS NULL THEN out_cost_sum = get_not_value(); END IF;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_man_hour_cost_sum IS 'Стоимость человекочаса по подрядчику и режиму в периоде';

CREATE FUNCTION get_man_hour_cost(InContractorId INTEGER, InModeId INTEGER, InTermId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS man_hour_cost_rates AS
$$
  SELECT *
  FROM man_hour_cost_rates
  WHERE mode_id = InModeId
  AND contractor_id = InContractorId
  AND term_id = InTermId;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_man_hour_cost IS 'Информация о стоимости человекочаса по подрядчику и режиму в периоде';

CREATE FUNCTION get_man_hour_cost_by_bldn_term(InBldnId INTEGER, InTermId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS man_hour_cost_rates AS
$$
  DECLARE
    _out_row man_hour_cost_rates%ROWTYPE;
    _contractor_id INTEGER;
  BEGIN

    WITH bldn_contractors AS (
      SELECT contractor_id FROM buildings_history WHERE bldn_id = InBldnId AND term_id = InTermId
       UNION ALL
      SELECT contractor_id FROM buildings WHERE id = InBldnId AND InTermId = (SELECT MAX(id) FROM terms)
    )
    SELECT contractor_id INTO _contractor_id FROM bldn_contractors;

    SELECT mr.* INTO _out_row
      FROM man_hour_cost_rates AS mr
      INNER JOIN bldn_man_hour_cost AS bc ON bc.term_id = mr.term_id AND bc.mode_id = mr.mode_id
     WHERE bldn_id = InBldnId
       AND contractor_id = _contractor_id
       AND mr.term_id = InTermId;

    RETURN _out_row;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_man_hour_cost_by_bldn_term IS 'Стоимость человекочаса в доме в периоде';

CREATE FUNCTION get_man_hour_cost_rate(InItemId BIGINT, InUserId INTEGER, InPCName VARCHAR) RETURNS man_hour_cost_rates AS
$$
  SELECT *
  FROM man_hour_cost_rates
  WHERE id = InItemId;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_man_hour_cost_rate IS 'Стоимости человекочаса по коду';

CREATE FUNCTION get_man_hour_cost_rates_by_term(InTermId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF man_hour_cost_rates AS
$$
  BEGIN
    IF NOT user_has_right_read(InUserId, rights_get_contractor_cost_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    RETURN QUERY
    SELECT *
      FROM man_hour_cost_rates
      WHERE term_id = InTermId
      ORDER BY contractor_id, mode_id;

    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_man_hour_cost_rates_by_term IS 'Стоимости человекочаса в периоде';
    
CREATE FUNCTION change_man_hour_cost_sum(InItemId BIGINT, InSum NUMERIC, InUserId INTEGER, InPCName VARCHAR) RETURNS VOID AS
$$
  DECLARE
    _cont_id INTEGER;
    _term_id INTEGER;
    _mode_id INTEGER;
  BEGIN
    IF NOT user_has_right_change(InUserId, rights_get_contractor_cost_rights_number()) THEN
      RAISE '%, %', get_error_number('has_no_access'), get_error_message('has_no_access');
    END IF;

    SELECT contractor_id, mode_id, term_id INTO _cont_id, _mode_id, _term_id
      FROM man_hour_cost_rates WHERE id = InItemId;
    
    WITH updated_rows AS (
      UPDATE man_hour_cost_rates
	 SET cost_sum = InSum
       WHERE id = InItemId
	     RETURNING *
    )
	INSERT INTO log_log(action_id, user_id, pc_name, log_action)
	SELECT 34, InUserId, InPCName, JSONB_AGG(ttt)
	  FROM (SELECT JSONB_AGG(mhr) AS prev, JSONB_AGG(updated_rows) AS upd
		  FROM man_hour_cost_rates AS mhr, updated_rows
		 WHERE mhr.id = InItemId) AS ttt;

    PERFORM recalc_maintenance_works(_cont_id, _term_id, _mode_id);

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION change_man_hour_cost_sum IS 'Изменение стоимости человекочаса';

CREATE FUNCTION get_subaccount_terms(InUserId INTEGER, InPCName VARCHAR) RETURNS SETOF terms AS
$$
  SELECT DISTINCT t.* FROM terms AS t
  INNER JOIN bldn_subaccounts AS bs ON bs.term_id = t.id
  ORDER BY t.begin_date DESC;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_subaccount_terms IS 'Периоды в которых есть информация о субсчетах';

CREATE FUNCTION get_occ_and_address(occ_list INTEGER[]) RETURNS TABLE (
  occ_id INTEGER
  , flat_no VARCHAR
  , address TEXT
  , fias VARCHAR
  , out_saldo NUMERIC
  )AS
$$
  DECLARE
    _max_rh_term INTEGER;
    _max_flats_term INTEGER;
  BEGIN
    
    SELECT MAX(term_id) INTO _max_rh_term FROM rkc_values_history;
    SELECT MAX(term_id) INTO _max_flats_term FROM flats;
    
    RETURN QUERY
      WITH occs AS (
	SELECT rh.occ_id, rh.flat_id, SUM(rh.out_saldo) AS out_saldo FROM rkc_values_history AS rh WHERE term_id = _max_rh_term AND rh.occ_id = ANY(occ_list) GROUP BY rh.occ_id, rh.flat_id
      ), fff AS (
	SELECT DISTINCT f.flat_id, f.bldn_id, f.flat_no FROM flats AS f WHERE term_id = _max_flats_term
      )
      SELECT occs.occ_id, fff.flat_no, bldn_address(b.id), b.fias, occs.out_saldo
      FROM occs
      INNER JOIN fff USING (flat_id)
      INNER JOIN buildings AS b ON fff.bldn_id = b.id;
    
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION get_occ_and_address IS 'Адреса и сальдо по списку лицевых счетов (для ответов собесу)';

CREATE FUNCTION get_employee_signature(InEmployeeId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS employees AS
$$
    SELECT * FROM employees
     WHERE id = InEmployeeId;
$$ LANGUAGE SQL STABLE;
COMMENT ON FUNCTION get_employee_signature IS 'Подпись работника';


-- BEGIN REPORTS

CREATE FUNCTION report_bldn_work_completition(InBldnId INTEGER, InTermId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS TABLE (
  out_work_name TEXT,
  out_work_sum NUMERIC,
  out_group INTEGER,
  out_priority INTEGER
) AS $$
BEGIN
  RETURN QUERY
    -- решить с текущим ремонтом? убрать
    -- СОИ
  SELECT name || CASE WHEN bt.has_thermoregulator AND expense_item = 5 THEN ' и узла автоматического регулирования системы теплоснабжения' ELSE '' END,
	 expense_fact_sum,
    group_id,
    report_priority
    FROM bldn_expenses AS be
    INNER JOIN buildings_tech_info AS bt ON bt.bldn_id = be.bldn_id
   WHERE be.bldn_id = InBldnId
     AND term_id = InTermId
    AND expense_item NOT IN (17, 10, 18, 19, 12)

    UNION

    SELECT wk.name,
    work_sum,
    2000000,
    0
    FROM works
    INNER JOIN work_kinds AS wk ON workkind_id = wk.id
    WHERE bldn_id = InBldnId
    AND work_date = InTermId
    AND gwt_id = 2
    AND finance_source = 1
   ORDER BY group_id, report_priority;

  RETURN;
END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION report_bldn_work_completition IS 'Акт выполенных работ дома';

CREATE FUNCTION report_year_plan(begin_year DATE) RETURNS
       TABLE ( out_bldn_id integer,
       	       out_address text,
	       out_work_name varchar,
	       out_contractor_name varchar,
	       out_month_name text,
	       out_work_sum numeric,
	       out_work_status varchar,
	       out_current_subaccount numeric,
	       out_plan_end_year numeric,
	       out_in_plan_flag integer,
	       out_plan_end_with_works numeric,
	       out_m1 numeric,
	       out_m2 numeric,
	       out_m3 numeric,
	       out_m4 numeric,
	       out_m5 numeric,
	       out_m6 numeric,
	       out_m7 numeric,
	       out_m8 numeric,
	       out_m9 numeric,
	       out_m10 numeric,
	       out_m11 numeric,
	       out_m12 numeric )AS $$
DECLARE
	_max_term INTEGER;
	_max_term_date DATE;
	_plan_months INTEGER;
	_begin_term INTEGER;
	_end_term INTEGER;
	_end_year DATE;
BEGIN
	_end_year := begin_year + INTERVAL '1 year - 1 day';
	SELECT MAX( term_id ) INTO _max_term FROM bldn_subaccounts;
	SELECT begin_date INTO _max_term_date FROM terms WHERE id = _max_term;
	SELECT id INTO _begin_term FROM terms WHERE begin_year BETWEEN begin_date AND end_date;
	SELECT id INTO _end_term FROM terms WHERE _end_year BETWEEN begin_date AND end_date;
	IF _end_term IS NULL THEN
	   SELECT MAX(id) INTO _end_term FROM terms;
	END IF;
	
	IF _max_term_date < begin_year
	THEN
		_plan_months := 12;
	ELSE
		_plan_months := 12 - EXTRACT( MONTH FROM _max_term_date );
	END IF;

	RETURN QUERY
	WITH tmp_works AS (
	     SELECT w.bldn_id,
	     	    w.id,
	       	    w.workkind_id,
		    wk.name AS wk_name,
		    w.contractor_id,
		    c.name AS cont_name,
		    t.begin_date AS wdate,
		    w.work_sum,
		    'Выполнена' AS status,
		    0 AS in_plan
	     FROM works w
       	     	  INNER JOIN terms t ON w.work_date = t.id
		  INNER JOIN work_kinds wk ON wk.id = w.workkind_id
		  INNER JOIN contractors c ON w.contractor_id = c.id
       	     WHERE w.work_date BETWEEN _begin_term AND _end_term
       	     	   AND w.gwt_id = 2

             UNION ALL

       	     SELECT w.bldn_id,
	     	    w.id,
       	      	    w.workkind_id,
		    wk.name,
	      	    w.contractor_id,
		    c.name,
	      	    w.work_date AS wdate,
		    COALESCE(NULLIF(smeta_sum, 0), w.work_sum),
		    pws.name AS status,
		    1
       	     FROM plan_works w
       	     	  INNER JOIN plan_work_statuses pws ON w.work_status = pws.id
		  INNER JOIN work_kinds wk ON w.workkind_id = wk.id
		  INNER JOIN contractors c ON c.id = w.contractor_id
       	     WHERE w.work_date BETWEEN begin_year AND _end_year
	     	   AND w.gwt_id = 2
       	     	   AND pws.in_plan
       ),
       plansubacc_sum AS (
       	    SELECT ps.bldn_id, TRUNC( (plan_sum * plan_percent * _plan_months )::NUMERIC, 2 ) AS plan_sum
	    FROM plan_subaccounts ps
       ),
       subacc_sum AS (
       	    SELECT w.bldn_id, SUM( w.work_sum ) AS new_works_sum
	    FROM works w
	    WHERE w.work_date > _max_term + 1
	    GROUP BY w.bldn_id
       )

	SELECT b.id,
	       xs.name || ' д.' || b.bldn_no AS address,
	       COALESCE( tw.wk_name, 'Накопление' ) AS work_name,
	       tw.cont_name,
	       to_char( tw.wdate, 'TMMonth' ) AS month_name,
	       tw.work_sum,
	       tw.status,
	       COALESCE( bs.subaccount_sum, 0 ) - COALESCE( ss.new_works_sum, 0 ) AS current_subaccount,
	       COALESCE( bs.subaccount_sum, 0 ) - COALESCE( ss.new_works_sum, 0 ) + ps.plan_sum AS end_year_sum,
	       tw.in_plan,
	       COALESCE( bs.subaccount_sum, 0 ) - COALESCE( ss.new_works_sum, 0 ) + ps.plan_sum - COALESCE( (SUM(tw.work_sum * tw.in_plan) OVER (PARTITION BY b.id ORDER BY tw.wdate, tw.wk_name, tw.in_plan, tw.id)), 0 ) AS sum_with_plan,
       	       CASE EXTRACT( MONTH FROM tw.wdate ) WHEN 1 THEN tw.work_sum ELSE 0 END AS m1,
       	       CASE EXTRACT( MONTH FROM tw.wdate ) WHEN 2 THEN tw.work_sum ELSE 0 END AS m2,
       	       CASE EXTRACT( MONTH FROM tw.wdate ) WHEN 3 THEN tw.work_sum ELSE 0 END AS m3,
       	       CASE EXTRACT( MONTH FROM tw.wdate ) WHEN 4 THEN tw.work_sum ELSE 0 END AS m4,
       	       CASE EXTRACT( MONTH FROM tw.wdate ) WHEN 5 THEN tw.work_sum ELSE 0 END AS m5,
       	       CASE EXTRACT( MONTH FROM tw.wdate ) WHEN 6 THEN tw.work_sum ELSE 0 END AS m6,
       	       CASE EXTRACT( MONTH FROM tw.wdate ) WHEN 7 THEN tw.work_sum ELSE 0 END AS m7,
       	       CASE EXTRACT( MONTH FROM tw.wdate ) WHEN 8 THEN tw.work_sum ELSE 0 END AS m8,
       	       CASE EXTRACT( MONTH FROM tw.wdate ) WHEN 9 THEN tw.work_sum ELSE 0 END AS m9,
       	       CASE EXTRACT( MONTH FROM tw.wdate ) WHEN 10 THEN tw.work_sum ELSE 0 END AS m10,
       	       CASE EXTRACT( MONTH FROM tw.wdate ) WHEN 11 THEN tw.work_sum ELSE 0 END AS m11,
       	       CASE EXTRACT( MONTH FROM tw.wdate ) WHEN 12 THEN tw.work_sum ELSE 0 END AS m12
	FROM managed_buildings b
	     INNER JOIN xstreets xs ON b.street_id = xs.id
	     LEFT JOIN bldn_subaccounts bs ON bs.bldn_id = b.id AND bs.term_id = _max_term
	     LEFT JOIN plansubacc_sum ps ON ps.bldn_id = b.id
	     LEFT JOIN subacc_sum ss ON b.id = ss.bldn_id
	     LEFT JOIN tmp_works tw ON tw.bldn_id = b.id
	ORDER BY xs.mid, xs.vid, address, tw.wdate, tw.wk_name, tw.in_plan;

END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION report_mainworkmaterials(InBeginDate INTEGER, InEndDate INTEGER, InContId INTEGER, InUserId INTEGER)
RETURNS TABLE(out_bldnid INTEGER, out_contractorname VARCHAR, out_address TEXT, out_workname TEXT, out_transport NUMERIC, out_materials NUMERIC, out_manhours NUMERIC, out_workdate INTEGER) AS
$$
DECLARE
	_work_round INTEGER;
BEGIN
	IF NOT has_access(InUserId, 'gwt', 1) THEN RAISE EXCEPTION SQLSTATE '60010' USING MESSAGE = 'Не хватает прав'; END IF;

	SELECT get_gwt_round() INTO _work_round;

	RETURN QUERY
	SELECT b.id,
	       c.name,
       	       xs.name || ' д.' || b.bldn_no AS address,
       	       wk.name || coalesce(' (' || w.note || ')', '') AS work_name,
	       ROUND( SUM( CASE WHEN wmt.is_transport THEN wm.material_count * wm.material_cost ELSE 0 END ), _work_round ) AS transport,
	       COALESCE( ROUND( SUM( CASE WHEN wmt.is_transport THEN 0 ELSE wm.material_count * wm.material_cost END ), _work_round ), 0) AS materials,
	       w.man_hours,
	       w.work_date
	FROM maintenance_works w
	     INNER JOIN work_kinds wk ON w.workkind_id = wk.id
	     INNER JOIN work_types wt ON wk.worktype_id = wt.id
	     INNER JOIN buildings b ON w.bldn_id = b.id
	     INNER JOIN xstreets xs ON b.street_id = xs.id
	     INNER JOIN contractors c ON w.contractor_id = c.id
	     LEFT JOIN works_materials wm ON w.id = wm.maintenance_work_id
	     LEFT JOIN work_material_types wmt ON wm.material_id = wmt.id
	WHERE  w.print_flag
	     AND w.work_date BETWEEN InBeginDate AND InEndDate
	     AND (c.id = InContId OR is_all_values(InContId))
	GROUP BY b.id, c.name, address, work_name, w.man_hours, w.work_date, c.id, w.id
	ORDER BY c.id, address, w.work_date;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION report_contractormaterials(InBeginDate INTEGER, InEndDate INTEGER, InContId INTEGER, InUserId INTEGER)
RETURNS TABLE(out_contractorname VARCHAR, out_materialname VARCHAR, out_materialsum NUMERIC, out_istransport BOOLEAN) AS
$$
BEGIN
	IF NOT has_access(InUserId, 'gwt', 1) THEN RAISE EXCEPTION SQLSTATE '60010' USING MESSAGE = 'Не хватает прав'; END IF;

	RETURN QUERY
	SELECT c.name,
	       wmt.material_name,
	       SUM( wm.material_count * wm.material_cost ),
	       wmt.is_transport
	FROM maintenance_works w
	     INNER JOIN contractors c ON w.contractor_id = c.id
	     INNER JOIN works_materials wm ON w.id = wm.maintenance_work_id
	     INNER JOIN work_material_types wmt ON wm.material_id = wmt.id
	WHERE  w.print_flag
	     AND w.work_date BETWEEN InBeginDate AND InEndDate
	     AND (c.id = InContId OR is_all_values(InContId))
	GROUP BY c.name, wmt.is_transport, wmt.material_name
	ORDER BY c.name, wmt.is_transport, wmt.material_name;
END;
$$ LANGUAGE plpgsql;

CREATE FUNCTION report_101 (InTermId INTEGER, InUkServiceId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS TABLE (OutBldnId INTEGER, OutAddress TEXT, OutSquare NUMERIC, OutPrice NUMERIC, OutAccrued NUMERIC, OutAddeds NUMERIC, OutAddedCom NUMERIC, OutAddedClean NUMERIC, OutAddedDolg NUMERIC, OutDiff NUMERIC, OutPaid NUMERIC, OutCompens NUMERIC) AS
$$
  BEGIN
    RETURN QUERY
    WITH prices AS (
      SELECT bldn_id,
	     SUM(price) AS prs
	FROM expenses e
	       INNER JOIN expense_items ei ON e.expense_item = ei.id
       WHERE (ei.uk_service_id = InUkServiceId or is_all_values(InUkServiceId))
	 AND term_id = InTermId
       GROUP BY bldn_id
    ),
      accrueds AS (
	SELECT bldn_id,
	       SUM(accrued) AS acc_sum,
	       SUM(added) AS add_sum,
	       SUM(paid) AS paid_sum,
	       SUM(compens) AS compens_sum
	  FROM rkc_values_history rh
		 INNER JOIN rkc_services rs ON rh.rkc_service_id = rs.id
	 WHERE term_id = InTermId
	   AND (rs.uk_service_id = InUkServiceId OR is_all_values(InUkServiceId))
	 GROUP BY bldn_id
      ),
      squares AS (
	SELECT
	  bldn_id
	  , SUM(square) AS total_square
	  FROM flats
	 WHERE term_id = InTermId
	 GROUP by bldn_id
      ),
      addeds AS (
	SELECT bldn_id
	       , SUM(CASE WHEN type_id = 1 THEN added_value ELSE 0 END) AS add1
	       , SUM(CASE WHEN type_id = 2 THEN added_value ELSE 0 END) AS add2
	       , SUM(CASE WHEN type_id = 3 THEN added_value ELSE 0 END) AS add3
	  FROM rkc_addeds_history
		 INNER JOIN rkc_services rs ON rs.id = service_id
	 WHERE term_id = InTermId
	   AND (rs.uk_service_id = InUkServiceId OR is_all_values(InUkServiceId))
	 GROUP BY bldn_id
      )
    SELECT b.id,
      xs.name1 || ' д. ' || b.bldn_no,
      s.total_square,
      prs,
      acc_sum,
      add_sum,
      ad.add1,
      ad.add2,
      ad.add3,
      ROUND( s.total_square * prs - acc_sum, 2) AS diff,
      paid_sum,
      compens_sum
      FROM prices p
      INNER JOIN accrueds a ON p.bldn_id = a.bldn_id
      INNER JOIN squares s ON s.bldn_id = p.bldn_id
      INNER JOIN buildings b ON p.bldn_id = b.id
      INNER JOIN xstreets xs ON b.street_id = xs.id
      LEFT JOIN addeds ad ON p.bldn_id = ad.bldn_id
     ORDER BY diff DESC, xs.mid, xs.vid, xs.name, b.bldn_no;

    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION report_101 IS 'Проверка начислений за указанный месяц';


CREATE FUNCTION report_101a (InBeginTermId INTEGER, InEndTermId INTEGER, InUkServiceId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS TABLE (OutBldnId INTEGER, OutAddress TEXT, OutSquare NUMERIC, OutPrice NUMERIC,OutAccrued NUMERIC, OutAddeds NUMERIC, OutAddedCom NUMERIC, OutAddedClean NUMERIC, OutAddedDolg NUMERIC, OutDiff NUMERIC, OutPaid NUMERIC, OutCompens NUMERIC) AS
$$
  BEGIN
    
    CREATE TEMPORARY TABLE _tmp_101a ON COMMIT DROP AS
      SELECT * FROM report_101(-1, InUkServiceId, InUserId, InPCName);
    
    FOR curTerm IN InBeginTermId..InEndTermId LOOP
      INSERT INTO _tmp_101a
      SELECT * FROM report_101(curTerm, InUkServiceId, InUserId, InPCName);
    END LOOP;

    RETURN QUERY
      SELECT t.OutBldnId,
      t.OutAddress,
      SUM (t.OutSquare),
      SUM (t.OutPrice),
      SUM (t.OutAccrued),
      SUM (t.OutAddeds),
      SUM (t.OutAddedCom),
      SUM (t.OutAddedClean),
      SUM (t.OutAddedDolg),
      SUM (t.OutDiff) AS diff,
      SUM (t.OutPaid),
      SUM (t.OutCompens)
      FROM _tmp_101a t
      GROUP BY t.OutBldnId, t.OutAddress
      ORDER BY diff DESC, OutAddress;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION report_101 IS 'Проверка начислений за указанный период';

CREATE FUNCTION report_102(InTypeId INTEGER, InBeginDate INTEGER, InEndDate INTEGER, InUserId INTEGER, InPCName VARCHAR)
  RETURNS TABLE (OutBldnId INTEGER
		 , OutAddress TEXT
		 , OutTermId INTEGER
		 , OutSum NUMERIC) AS
$$
  BEGIN
    RETURN QUERY
      SELECT bldn_id
      , bldn_address(bldn_id) AS address
      , term_id
      , sum(added_value)
      FROM rkc_addeds_history
      WHERE term_id BETWEEN InBeginDate AND InEndDate
      AND type_id = InTypeId
      GROUP by bldn_id, address, term_id
      ORDER BY address, term_id DESC;

    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;

CREATE FUNCTION report_11(InMCId INTEGER, InMDId INTEGER, InVillageId INTEGER, InContractorId INTEGER, InDate INTEGER, InUserId INTEGER, InPCName VARCHAR)
RETURNS TABLE (OutBldnId INTEGER
	       , OutTermId INTEGER
	       , OutAddress TEXT
	       , OutSum NUMERIC
	       , OutPlanSum NUMERIC
	       , OutPercent NUMERIC
	       , OutBeginValue NUMERIC
	       , OutAccrued NUMERIC
	       , OutPaid NUMERIC
	       ) AS
$$
  DECLARE
    _max_sa_date DATE;
    _year_begin_term INTEGER;
    _use_all_village BOOLEAN;
    _use_all_contractors BOOLEAN;
    _use_all_md BOOLEAN;
    _use_all_mc BOOLEAN;
  BEGIN
    SELECT begin_date INTO _max_sa_date
    FROM terms t
      INNER JOIN bldn_subaccounts bs ON t.id = bs.term_id AND bs.term_id = InDate;

    IF _max_sa_date IS NULL THEN
      RAISE '%, %', get_error_number('has_no_values'), ger_error_message('has_no_values');
    END IF;
    
    SELECT is_all_values(InVillageId),
	   is_all_values(InContractorId),
	   is_all_values(InMDId),
	   is_all_values(InMCId)
      INTO _use_all_village,
	   _use_all_contractors,
	   _use_all_md,
	   _use_all_mc;

    SELECT MIN(term_id) INTO _year_begin_term
      FROM bldn_subaccounts sa
	   INNER JOIN terms t ON sa.term_id = t.id
     WHERE EXTRACT(YEAR FROM t.begin_date) = EXTRACT(YEAR FROM _max_sa_date);

    RETURN QUERY
    WITH begin_sa_values AS (
      SELECT bldn_id,
	     subaccount_sum
	FROM bldn_subaccounts
       WHERE term_id = _year_begin_term
    ),
      used_buildings AS (
	SELECT bldn_id,
	       bldn_no,
	       street_id
	  FROM buildings_history
	 WHERE term_id = InDate
	   AND dogovor_type > 0
	   AND (mc_id = InMCId OR _use_all_mc)
	   AND (contractor_id = InContractorId OR _use_all_contractors)
      ),
      sa_year_sum AS (
	SELECT bldn_id, SUM(accrued_sum) AS accrued, SUM(paid_sum) AS paid
	  FROM sub_accounts
	       INNER JOIN terms t on term_id = t.id
	 WHERE t.begin_date BETWEEN (SELECT begin_date FROM terms WHERE id = _year_begin_term) AND _max_sa_date
	 GROUP BY bldn_id
      )
      SELECT bs.bldn_id
      , t.id
      , s.name || ' д.' || bldn_no
      , bs.subaccount_sum
      , bp.plan_sum
      , bldn_subaccount_percent(bs.bldn_id, t.begin_date, InUserId, InPCName)
      , bv.subaccount_sum
      , say.accrued
      , say.paid
      FROM bldn_subaccounts bs
      INNER JOIN used_buildings mb USING (bldn_id)
      INNER JOIN begin_sa_values bv USING (bldn_id)
      INNER JOIN sa_year_sum say USING (bldn_id)
      INNER JOIN xstreets s ON mb.street_id = s.id
      INNER JOIN plan_subaccounts bp USING (bldn_id)
      INNER JOIN terms t ON bs.term_id = t.id
     WHERE t.begin_date = _max_sa_date
      AND (s.vid = InVillageId OR _use_all_village)
      AND (s.mid = InMDId OR _use_all_md)
      ORDER BY s.mid, 3;
    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION report_1<1 IS 'Текущее состояние субсчетов';

CREATE FUNCTION report_12(InBldnId INTEGER, InBeginTermId INTEGER, InEndTermId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS TABLE (
  OutOccId INTEGER,
  OutFlat TEXT,
  OutFIO TEXT,
  OutAccrued NUMERIC,
  OutAddeds NUMERIC,
  OutCompens NUMERIC,
  OutPaids NUMERIC,
  OutAccruedCount BIGINT,
  OutPaidCount BIGINT,
  OutInSaldo NUMERIC,
  OutOutSaldo NUMERIC
) AS
$$
  DECLARE
    _begin_date DATE;
    _end_date DATE;    
  BEGIN
    SELECT begin_date INTO _begin_date FROM terms WHERE id = InBeginTermId;
    SELECT begin_date INTO _end_date FROM terms WHERE id = InEndTermId;

    RETURN QUERY
    WITH paid_sum AS (
      SELECT
	occ_id,
	SUM(paid) AS paids,
	COUNT(DISTINCT(term_id)) AS count_paids
	FROM rkc_values_history
	     INNER JOIN terms AS t ON term_id = t.id
       WHERE t.begin_date BETWEEN (_begin_date + INTERVAL '1 month') AND (_end_date + INTERVAL '1 month')
	 AND bldn_id = InBldnId
       GROUP BY occ_id
    ), accrued_sum AS (
      SELECT
	occ_id,
	STRING_AGG(DISTINCT(flat_no), ', ') AS flat,
	SUM(accrued) AS accrued,
	SUM(added) AS added,
	SUM(compens) AS compens,
	COUNT(DISTINCT(term_id)) AS count_accrued
	FROM rkc_values_history
	       INNER JOIN terms AS t ON term_id = t.id
       WHERE t.begin_date BETWEEN _begin_date AND _end_date
	 AND bldn_id = InBldnId
       GROUP BY occ_id
    ), in_saldo AS (
      SELECT occ_id, SUM(in_saldo) AS in_saldo
	FROM rkc_values_history
       WHERE term_id = InBeginTermId
	     AND bldn_id = InBldnId
       GROUP BY occ_id
    ), out_saldo AS (
      SELECT occ_id, SUM(out_saldo) AS out_saldo
	FROM rkc_values_history
       WHERE term_id = InEndTermId
	     AND bldn_id = InBldnId
       GROUP BY occ_id
    ), fios AS (
      SELECT occ_id, STRING_AGG(DISTINCT(owner_name), ', ') AS fios
	FROM owners AS o
	     INNER JOIN flat_shares AS fls ON o.share_id = fls.id
	     INNER JOIN flats AS fh USING (flat_id, term_id)
	     INNER JOIN rkc_values_history AS rh USING ( flat_id, term_id, bldn_id)
       WHERE term_id = (SELECT MAX(term_id) FROM rkc_values_history WHERE term_id BETWEEN InBeginTermId AND InEndTermId)
	     AND fh.bldn_id = InBldnId
       GROUP BY occ_id
    )
    SELECT
      acs.occ_id,
      acs.flat,
      fios.fios,
      acs.accrued,
      acs.added,
      acs.compens,
      ps.paids,
      acs.count_accrued,
      ps.count_paids,
      b.in_saldo,
      e.out_saldo
      FROM accrued_sum AS acs
      INNER JOIN paid_sum AS ps ON acs.occ_id = ps.occ_id
      LEFT JOIN in_saldo AS b ON b.occ_id = acs.occ_id
      LEFT JOIN out_saldo AS e ON e.occ_id = acs.occ_id
      Left JOIN fios ON fios.occ_id = acs.occ_id
      ORDER BY acs.flat;

    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION report_12 IS 'Собираемость за период по выбранному дому';

CREATE FUNCTION report_13(InBeginTerm INTEGER, InEndTerm INTEGER, InUserId INTEGER, InPCName VARCHAR)
  RETURNS TABLE (OutBldnId INTEGER, OutAddress TEXT, OutTerm DATE, OutService VARCHAR, OutSum NUMERIC) AS
$$
  DECLARE
    _begin_date DATE;
    _end_date DATE;
  BEGIN
    SELECT begin_date INTO _begin_date FROM terms WHERE id = InBeginTerm;
    SELECT end_date INTO _end_date FROM terms WHERE id = InEndTerm;

    RETURN QUERY
    SELECT rvh.bldn_id AS OutBldnId
	   , bldn_address(rvh.bldn_id) AS OutAddress
	   , terms.begin_date AS OutTerm
	   , rs.full_name AS OutService
	   , SUM(rvh.paid) AS OutSum
      FROM rkc_values_history AS rvh
	   JOIN buildings_history AS bh ON bh.bldn_id = rvh.bldn_id AND bh.term_id + 1 = rvh.term_id
	   JOIN rkc_services AS rs ON rvh.rkc_service_id = rs.id
	   JOIN terms ON rvh.term_id = terms.id
     WHERE bh.dogovor_type = 0
       AND terms.begin_date BETWEEN _begin_date AND _end_date
     GROUP BY rvh.bldn_id, terms.begin_date, rs.full_name
      HAVING SUM(rvh.paid) != 0
     ORDER BY OutAddress, OutTerm;

    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION report_13 IS 'Оплата по ушедшим домам';

CREATE FUNCTION report_14(InBeginTermId INTEGER, InEndTermId INTEGER, InServiceId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS TABLE (
  outBldnId INTEGER,
  outAddress TEXT,
  outAccrued NUMERIC,
  outCompens NUMERIC,
  outFullAddeds NUMERIC,
  outDolgAddeds NUMERIC,
  outPaid NUMERIC
) AS $$
  BEGIN
    RETURN QUERY
    WITH acc AS (
      SELECT bldn_id , SUM(accrued) AS accrued, SUM(compens) AS compens, SUM(added) AS added
	FROM rkc_values_history AS rh
	     INNER JOIN rkc_services AS rs ON rh.rkc_service_id = rs.id
       WHERE term_id BETWEEN InBeginTermId AND InEndTermId
	 AND rs.total_uk_service_id = InServiceId
       GROUP BY bldn_id
    ), paids AS (
      SELECT bldn_id, SUM(paid) AS paid
	FROM rkc_values_history AS rh
	     INNER JOIN rkc_services AS rs ON rh.rkc_service_id = rs.id
       WHERE term_id BETWEEN InBeginTermId + 1 AND InEndTermId + 1
	 AND rs.total_uk_service_id = InServiceId
       GROUP BY bldn_id
    ), addeds AS (
      SELECT bldn_id,
	     SUM(added_value) AS advalue
	FROM rkc_addeds_history AS rh
	     INNER JOIN rkc_services AS rs ON rh.service_id = rs.id
       WHERE term_id BETWEEN InBeginTermId AND InEndTermId
	 AND rs.total_uk_service_id = InServiceId
	 AND type_id = 3
       GROUP BY bldn_id
    )
    SELECT acc.bldn_id,
      s.name || ' д.' || bldn_no,
	   acc.accrued,
	   acc.compens,
	   acc.added,
           COALESCE(addeds.advalue, 0),
	   paids.paid
      FROM acc
      INNER JOIN buildings AS b ON acc.bldn_id = b.id
      INNER JOIN xstreets AS s ON b.street_id = s.id
	   FULL JOIN paids USING (bldn_id)
	   LEFT JOIN addeds USING (bldn_id)
     ORDER BY s.mid, s.vid, s.name, b.bldn_no;
    RETURN;
  END;
$$ LANGUAGE plpgsql;
COMMENT ON FUNCTION report_14 IS 'Собираемость за период по домам';

CREATE FUNCTION report_201 (InBeginTerm INTEGER, InEndTerm INTEGER, InGwtId INTEGER, InUserId INTEGER, InPCName VARCHAR) RETURNS TABLE (
  OutBldnId INTEGER
  , OutMDName VARCHAR
  , OutAddress TEXT
  , OutSquare NUMERIC
  , OutWorkSum NUMERIC
  , OutFlatTerm INTEGER
) AS
$$
  DECLARE
    _begin_date DATE;
    _end_date DATE;
    _flat_term INTEGER;

  BEGIN

    SELECT begin_date INTO _begin_date FROM terms WHERE id = InBeginTerm;
    SELECT begin_date INTO _end_date FROM terms WHERE id = InEndTerm;
    SELECT MAX(term_id) INTO _flat_term FROM flats WHERE term_id BETWEEN InBeginTerm AND InEndTerm;
    IF _flat_term IS NULL then
      SELECT MAX(term_id) INTO _flat_term FROM flats;
    END IF;

    RETURN QUERY
      WITH sq AS (
	SELECT
	  bldn_id
	  , _flat_term AS flat_term
	  , SUM(square) AS fsquare
	  FROM flats
	 WHERE term_id = _flat_term
	 GROUP BY bldn_id
      )
    SELECT
      b.id INTEGER
      , md.name
      , xs.name1 || ' д. ' || b.bldn_no
      , f.fsquare
      , SUM(w.work_sum)
      , f.flat_term
      FROM buildings AS b
      INNER JOIN xstreets AS xs ON b.street_id = xs.id
      INNER JOIN municipal_districts AS md ON xs.mid = md.id
      INNER JOIN sq AS f ON f.bldn_id = b.id
      INNER JOIN works AS w ON w.bldn_id = b.id
      INNER JOIN terms ON terms.id = w.work_date AND w.gwt_id = InGwtId
     WHERE terms.begin_date BETWEEN _begin_date AND _end_date
      GROUP BY b.id, md.name, xs.name, xs.name1, b.bldn_no, f.fsquare, f.flat_term
      ORDER BY md.name, xs.name, b.bldn_no;

    RETURN;
  END;
$$ LANGUAGE plpgsql STABLE;
COMMENT ON FUNCTION report_201 IS '22-жкх зима';
