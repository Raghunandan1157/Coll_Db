-- =====================================================================
-- Employee Master Hierarchy Restructure — April 26, 2026
-- Source: 'New Area and Division ReStructure April 26 (1).xlsx'
-- Generated: 2026-04-16
--
-- Normalizes area_name/division_name/region_name/area_manager/division_manager
-- in employee_master for all 127 branches to match the new structure.
-- Fixes split-hierarchy issues (e.g., Hospet branch rows with both Davanagere
-- and Hospet area) by forcing all rows per UPPER(branch_name) to the new value.
--
-- Casing conventions:
--   - Region: Title Case, with 'AP' for Andra Pradesh and 'TS' for Telangana
--   - Division: Title Case, with 'AP' / 'TS' overrides (matches existing style)
--   - Area: Title Case
--   - AM / DM names: Title Case
-- =====================================================================

BEGIN;

UPDATE employee_master SET area_name='Kalburgi', division_name='Kalburgi', region_name='Kalburgi', area_manager='Prashanth', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'AFZALPUR';
UPDATE employee_master SET area_name='Kadur', division_name='Chitradurga', region_name='Chitradurga', area_manager='Kumar R', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'AJJAMPURA';
UPDATE employee_master SET area_name='Kalburgi', division_name='Kalburgi', region_name='Kalburgi', area_manager='Prashanth', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'ALAND';
UPDATE employee_master SET area_name='Indi', division_name='Kalburgi', region_name='Kalburgi', area_manager='Rajkumar Pawar', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'ALMEL';
UPDATE employee_master SET area_name='Chikkodi', division_name='Belagavi', region_name='Dharwad', area_manager='Yalagouda Khoth', division_manager='Sunil Kuber Malage' WHERE UPPER(branch_name) = 'ATHANI';
UPDATE employee_master SET area_name='Bidar', division_name='Bidar', region_name='Kalburgi', area_manager='Manohar', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'AURAD';
UPDATE employee_master SET area_name='Badami', division_name='Hubli', region_name='Dharwad', area_manager='Basavarajappa', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'BADAMI';
UPDATE employee_master SET area_name='Bagalkot', division_name='Belagavi', region_name='Dharwad', area_manager='Sharanabasappa', division_manager='Sunil Kuber Malage' WHERE UPPER(branch_name) = 'BAGALKOT';
UPDATE employee_master SET area_name='Chikballapura', division_name='Doddaballapura', region_name='Tumkur', area_manager='Ravi', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'BAGEPALLI';
UPDATE employee_master SET area_name='Belagavi', division_name='Belagavi', region_name='Dharwad', area_manager='Hajiali Tigadi', division_manager='Sunil Kuber Malage' WHERE UPPER(branch_name) = 'BAILHONGAL';
UPDATE employee_master SET area_name='Ballari', division_name='Hospet', region_name='Chitradurga', area_manager='Bairappa', division_manager='Ambarish M S' WHERE UPPER(branch_name) = 'BALLARI';
UPDATE employee_master SET area_name='Kolar', division_name='Doddaballapura', region_name='Tumkur', area_manager='Shivaraja N', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'BANGARPET';
UPDATE employee_master SET area_name='Humnabad', division_name='Bidar', region_name='Kalburgi', area_manager='Mallikarjun Waghraj', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'BASAVAKALYAN';
UPDATE employee_master SET area_name='Belagavi', division_name='Belagavi', region_name='Dharwad', area_manager='Hajiali Tigadi', division_manager='Sunil Kuber Malage' WHERE UPPER(branch_name) = 'BELAGAVI';
UPDATE employee_master SET area_name='Kolar', division_name='Doddaballapura', region_name='Tumkur', area_manager='Shivaraja N', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'BETHAMANGALA';
UPDATE employee_master SET area_name='Bidar', division_name='Bidar', region_name='Kalburgi', area_manager='Manohar', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'BHALKI';
UPDATE employee_master SET area_name='Bidar', division_name='Bidar', region_name='Kalburgi', area_manager='Manohar', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'BIDAR';
UPDATE employee_master SET area_name='Bidar', division_name='Bidar', region_name='Kalburgi', area_manager='Manohar', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'BIDAR-2';
UPDATE employee_master SET area_name='Bagalkot', division_name='Belagavi', region_name='Dharwad', area_manager='Sharanabasappa', division_manager='Sunil Kuber Malage' WHERE UPPER(branch_name) = 'BILAGI';
UPDATE employee_master SET area_name='Kadapa', division_name='AP', region_name='AP', area_manager='Putta Prasad', division_manager='A V Chiranjeevi' WHERE UPPER(branch_name) = 'BUDWAL';
UPDATE employee_master SET area_name='Indi', division_name='Kalburgi', region_name='Kalburgi', area_manager='Rajkumar Pawar', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'CHADCHAN';
UPDATE employee_master SET area_name='Chitradurga', division_name='Chitradurga', region_name='Chitradurga', area_manager='Manoj Naik B', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'CHALLAKERE';
UPDATE employee_master SET area_name='Bangalore Urban', division_name='Doddaballapura', region_name='Tumkur', area_manager='Mubarak S', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'CHANDAPURA';
UPDATE employee_master SET area_name='Chitradurga', division_name='Chitradurga', region_name='Chitradurga', area_manager='Manoj Naik B', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'CHANNAGIRI';
UPDATE employee_master SET area_name='Chikballapura', division_name='Doddaballapura', region_name='Tumkur', area_manager='Ravi', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'CHIKBALLAPURA';
UPDATE employee_master SET area_name='Chikkamagaluru', division_name='Tumkur', region_name='Tumkur', area_manager='Anil Kumar', division_manager='Mahesh' WHERE UPPER(branch_name) = 'CHIKKAMAGALURU';
UPDATE employee_master SET area_name='Tiptur', division_name='Tumkur', region_name='Tumkur', area_manager='P T Venkatesh', division_manager='Mahesh' WHERE UPPER(branch_name) = 'CHIKKANAYAKANAHALLI';
UPDATE employee_master SET area_name='Chikkodi', division_name='Belagavi', region_name='Dharwad', area_manager='Yalagouda Khoth', division_manager='Sunil Kuber Malage' WHERE UPPER(branch_name) = 'CHIKKODI';
UPDATE employee_master SET area_name='Sedam', division_name='Bidar', region_name='Kalburgi', area_manager='Mallikarjun', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'CHINCHOLI';
UPDATE employee_master SET area_name='Chikballapura', division_name='Doddaballapura', region_name='Tumkur', area_manager='Ravi', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'CHINTAMANI';
UPDATE employee_master SET area_name='Chitradurga', division_name='Chitradurga', region_name='Chitradurga', area_manager='Manoj Naik B', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'CHITRADURGA';
UPDATE employee_master SET area_name='Tumkur', division_name='Tumkur', region_name='Tumkur', area_manager='Dilip Kumar', division_manager='Mahesh' WHERE UPPER(branch_name) = 'DABUSPET';
UPDATE employee_master SET area_name='Davanagere', division_name='Chitradurga', region_name='Chitradurga', area_manager='Kotresha M', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'DAVANAGERE';
UPDATE employee_master SET area_name='Shahapur', division_name='Kalburgi', region_name='Kalburgi', area_manager='Basavaraj', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'DEVADURGA';
UPDATE employee_master SET area_name='Doddaballapura', division_name='Doddaballapura', region_name='Tumkur', area_manager='Prabhu', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'DEVANAHALLI';
UPDATE employee_master SET area_name='Kadapa', division_name='AP', region_name='AP', area_manager='Putta Prasad', division_manager='A V Chiranjeevi' WHERE UPPER(branch_name) = 'DHARMAVARAM';
UPDATE employee_master SET area_name='Dharwad', division_name='Hubli', region_name='Dharwad', area_manager='Sagar Sureban', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'DHARWAD';
UPDATE employee_master SET area_name='Doddaballapura', division_name='Doddaballapura', region_name='Tumkur', area_manager='Prabhu', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'DODDABALLAPURA';
UPDATE employee_master SET area_name='Gadag', division_name='Hubli', region_name='Dharwad', area_manager='Chandragoud Patil', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'GADAG';
UPDATE employee_master SET area_name='Mahaboobnagar', division_name='TS', region_name='TS', area_manager='P Maddilety', division_manager='A V Chiranjeevi' WHERE UPPER(branch_name) = 'GADWAL';
UPDATE employee_master SET area_name='Badami', division_name='Hubli', region_name='Dharwad', area_manager='Basavarajappa', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'GAJENDRAGAD';
UPDATE employee_master SET area_name='Kushtagi', division_name='Hubli', region_name='Dharwad', area_manager='Virupakshappa Choudi', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'GANGAVATHI';
UPDATE employee_master SET area_name='Belagavi', division_name='Belagavi', region_name='Dharwad', area_manager='Hajiali Tigadi', division_manager='Sunil Kuber Malage' WHERE UPPER(branch_name) = 'GOKAK';
UPDATE employee_master SET area_name='Doddaballapura', division_name='Doddaballapura', region_name='Tumkur', area_manager='Prabhu', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'GOWRIBIDANUR';
UPDATE employee_master SET area_name='Tiptur', division_name='Tumkur', region_name='Tumkur', area_manager='P T Venkatesh', division_manager='Mahesh' WHERE UPPER(branch_name) = 'GUBBI';
UPDATE employee_master SET area_name='Hospet', division_name='Hospet', region_name='Chitradurga', area_manager='Pavankumar Sherkhane', division_manager='Ambarish M S' WHERE UPPER(branch_name) = 'HAGARIBOMMANAHALLI';
UPDATE employee_master SET area_name='Kotturu', division_name='Hospet', region_name='Chitradurga', area_manager='A B Manjunatha', division_manager='Ambarish M S' WHERE UPPER(branch_name) = 'HARAPANAHALLI';
UPDATE employee_master SET area_name='Davanagere', division_name='Chitradurga', region_name='Chitradurga', area_manager='Kotresha M', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'HARIHARA';
UPDATE employee_master SET area_name='Bangalore Urban', division_name='Doddaballapura', region_name='Tumkur', area_manager='Mubarak S', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'HEBBAL';
UPDATE employee_master SET area_name='Chitradurga', division_name='Chitradurga', region_name='Chitradurga', area_manager='Manoj Naik B', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'HIRIYUR';
UPDATE employee_master SET area_name='Chitradurga', division_name='Chitradurga', region_name='Chitradurga', area_manager='Manoj Naik B', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'HOLAKERE';
UPDATE employee_master SET area_name='Davanagere', division_name='Chitradurga', region_name='Chitradurga', area_manager='Kotresha M', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'HONNALI';
UPDATE employee_master SET area_name='Chitradurga', division_name='Chitradurga', region_name='Chitradurga', area_manager='Manoj Naik B', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'HOSADURGA';
UPDATE employee_master SET area_name='Hospet', division_name='Hospet', region_name='Chitradurga', area_manager='Pavankumar Sherkhane', division_manager='Ambarish M S' WHERE UPPER(branch_name) = 'HOSPET';
UPDATE employee_master SET area_name='Dharwad', division_name='Hubli', region_name='Dharwad', area_manager='Sagar Sureban', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'HUBLI';
UPDATE employee_master SET area_name='Dharwad', division_name='Hubli', region_name='Dharwad', area_manager='Sagar Sureban', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'HUBLI-2';
UPDATE employee_master SET area_name='Tiptur', division_name='Tumkur', region_name='Tumkur', area_manager='P T Venkatesh', division_manager='Mahesh' WHERE UPPER(branch_name) = 'HULIYAR';
UPDATE employee_master SET area_name='Humnabad', division_name='Bidar', region_name='Kalburgi', area_manager='Mallikarjun Waghraj', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'HULSOOR';
UPDATE employee_master SET area_name='Humnabad', division_name='Bidar', region_name='Kalburgi', area_manager='Mallikarjun Waghraj', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'HUMNABAD';
UPDATE employee_master SET area_name='Kushtagi', division_name='Hubli', region_name='Dharwad', area_manager='Virupakshappa Choudi', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'HUNGUND';
UPDATE employee_master SET area_name='Hospet', division_name='Hospet', region_name='Chitradurga', area_manager='Pavankumar Sherkhane', division_manager='Ambarish M S' WHERE UPPER(branch_name) = 'HUVENAHADAGALLI';
UPDATE employee_master SET area_name='Indi', division_name='Kalburgi', region_name='Kalburgi', area_manager='Rajkumar Pawar', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'INDI';
UPDATE employee_master SET area_name='Bangalore Urban', division_name='Doddaballapura', region_name='Tumkur', area_manager='Mubarak S', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'J P NAGAR';
UPDATE employee_master SET area_name='Kotturu', division_name='Hospet', region_name='Chitradurga', area_manager='A B Manjunatha', division_manager='Ambarish M S' WHERE UPPER(branch_name) = 'JAGALORE';
UPDATE employee_master SET area_name='Bagalkot', division_name='Belagavi', region_name='Dharwad', area_manager='Sharanabasappa', division_manager='Sunil Kuber Malage' WHERE UPPER(branch_name) = 'JAMAKHANDI';
UPDATE employee_master SET area_name='Kalburgi', division_name='Kalburgi', region_name='Kalburgi', area_manager='Prashanth', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'JEVARGI';
UPDATE employee_master SET area_name='Kadapa', division_name='AP', region_name='AP', area_manager='Putta Prasad', division_manager='A V Chiranjeevi' WHERE UPPER(branch_name) = 'KADAPA';
UPDATE employee_master SET area_name='Kadapa', division_name='AP', region_name='AP', area_manager='Putta Prasad', division_manager='A V Chiranjeevi' WHERE UPPER(branch_name) = 'KADIRI';
UPDATE employee_master SET area_name='Kadur', division_name='Chitradurga', region_name='Chitradurga', area_manager='Kumar R', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'KADUR';
UPDATE employee_master SET area_name='Kalburgi', division_name='Kalburgi', region_name='Kalburgi', area_manager='Prashanth', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'KALABURAGI';
UPDATE employee_master SET area_name='Sedam', division_name='Bidar', region_name='Kalburgi', area_manager='Mallikarjun', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'KALAGI';
UPDATE employee_master SET area_name='Kalburgi', division_name='Kalburgi', region_name='Kalburgi', area_manager='Prashanth', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'KALBURGI-2';
UPDATE employee_master SET area_name='Dharwad', division_name='Hubli', region_name='Dharwad', area_manager='Sagar Sureban', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'KALGHATGI';
UPDATE employee_master SET area_name='Humnabad', division_name='Bidar', region_name='Kalburgi', area_manager='Mallikarjun Waghraj', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'KAMALAPURA';
UPDATE employee_master SET area_name='Bangalore Urban', division_name='Doddaballapura', region_name='Tumkur', area_manager='Mubarak S', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'KENGERI';
UPDATE employee_master SET area_name='Kotturu', division_name='Hospet', region_name='Chitradurga', area_manager='A B Manjunatha', division_manager='Ambarish M S' WHERE UPPER(branch_name) = 'KHANAHOSAHALLI';
UPDATE employee_master SET area_name='Belagavi', division_name='Belagavi', region_name='Dharwad', area_manager='Hajiali Tigadi', division_manager='Sunil Kuber Malage' WHERE UPPER(branch_name) = 'KITTUR';
UPDATE employee_master SET area_name='Sangareddy', division_name='TS', region_name='TS', area_manager='Sampath Kumar', division_manager='A V Chiranjeevi' WHERE UPPER(branch_name) IN ('KODANGAL','KODANGAL(VIKARABAD)');
UPDATE employee_master SET area_name='Kolar', division_name='Doddaballapura', region_name='Tumkur', area_manager='Shivaraja N', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'KOLAR';
UPDATE employee_master SET area_name='Kushtagi', division_name='Hubli', region_name='Dharwad', area_manager='Virupakshappa Choudi', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'KOPPAL';
UPDATE employee_master SET area_name='Tumkur', division_name='Tumkur', region_name='Tumkur', area_manager='Dilip Kumar', division_manager='Mahesh' WHERE UPPER(branch_name) = 'KORATAGERE';
UPDATE employee_master SET area_name='Kotturu', division_name='Hospet', region_name='Chitradurga', area_manager='A B Manjunatha', division_manager='Ambarish M S' WHERE UPPER(branch_name) = 'KOTTURU';
UPDATE employee_master SET area_name='Ballari', division_name='Hospet', region_name='Chitradurga', area_manager='Bairappa', division_manager='Ambarish M S' WHERE UPPER(branch_name) = 'KUDATHINI';
UPDATE employee_master SET area_name='Hospet', division_name='Hospet', region_name='Chitradurga', area_manager='Pavankumar Sherkhane', division_manager='Ambarish M S' WHERE UPPER(branch_name) = 'KUDLIGI';
UPDATE employee_master SET area_name='Tumkur', division_name='Tumkur', region_name='Tumkur', area_manager='Dilip Kumar', division_manager='Mahesh' WHERE UPPER(branch_name) = 'KUNIGAL';
UPDATE employee_master SET area_name='Kushtagi', division_name='Hubli', region_name='Dharwad', area_manager='Virupakshappa Choudi', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'KUSHTAGI';
UPDATE employee_master SET area_name='Gadag', division_name='Hubli', region_name='Dharwad', area_manager='Chandragoud Patil', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'LAXMESHWAR';
UPDATE employee_master SET area_name='Lingsugur', division_name='Bidar', region_name='Kalburgi', area_manager='Nagaraj Itagi', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'LINGSUGUR';
UPDATE employee_master SET area_name='Bagalkot', division_name='Belagavi', region_name='Dharwad', area_manager='Sharanabasappa', division_manager='Sunil Kuber Malage' WHERE UPPER(branch_name) = 'LOKAPUR';
UPDATE employee_master SET area_name='Tumkur', division_name='Tumkur', region_name='Tumkur', area_manager='Dilip Kumar', division_manager='Mahesh' WHERE UPPER(branch_name) = 'MADHUGIRI';
UPDATE employee_master SET area_name='Mahaboobnagar', division_name='TS', region_name='TS', area_manager='P Maddilety', division_manager='A V Chiranjeevi' WHERE UPPER(branch_name) = 'MAHABUB NAGAR';
UPDATE employee_master SET area_name='Kolar', division_name='Doddaballapura', region_name='Tumkur', area_manager='Shivaraja N', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'MALUR';
UPDATE employee_master SET area_name='Lingsugur', division_name='Bidar', region_name='Kalburgi', area_manager='Nagaraj Itagi', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'MANVI';
UPDATE employee_master SET area_name='Mahaboobnagar', division_name='TS', region_name='TS', area_manager='P Maddilety', division_manager='A V Chiranjeevi' WHERE UPPER(branch_name) = 'MARIKAL';
UPDATE employee_master SET area_name='Chikkodi', division_name='Belagavi', region_name='Dharwad', area_manager='Yalagouda Khoth', division_manager='Sunil Kuber Malage' WHERE UPPER(branch_name) = 'MUDALAGI';
UPDATE employee_master SET area_name='Vijayapur', division_name='Kalburgi', region_name='Kalburgi', area_manager='Prakash Biradar', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'MUDDEBIHAL';
UPDATE employee_master SET area_name='Chikkamagaluru', division_name='Tumkur', region_name='Tumkur', area_manager='Anil Kumar', division_manager='Mahesh' WHERE UPPER(branch_name) = 'MUDIGERE';
UPDATE employee_master SET area_name='Gadag', division_name='Hubli', region_name='Dharwad', area_manager='Chandragoud Patil', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'MUNDARAGI';
UPDATE employee_master SET area_name='Badami', division_name='Hubli', region_name='Dharwad', area_manager='Basavarajappa', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'NARAGUNDA';
UPDATE employee_master SET area_name='Sangareddy', division_name='TS', region_name='TS', area_manager='Sampath Kumar', division_manager='A V Chiranjeevi' WHERE UPPER(branch_name) = 'NARAYANKHED';
UPDATE employee_master SET area_name='Chikkodi', division_name='Belagavi', region_name='Dharwad', area_manager='Yalagouda Khoth', division_manager='Sunil Kuber Malage' WHERE UPPER(branch_name) = 'NIPPANI';
UPDATE employee_master SET area_name='Chikkamagaluru', division_name='Tumkur', region_name='Tumkur', area_manager='Anil Kumar', division_manager='Mahesh' WHERE UPPER(branch_name) = 'NR PURA';
UPDATE employee_master SET area_name='Kadur', division_name='Chitradurga', region_name='Chitradurga', area_manager='Kumar R', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'PANCHANHALLI';
UPDATE employee_master SET area_name='Lingsugur', division_name='Bidar', region_name='Kalburgi', area_manager='Nagaraj Itagi', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'RAICHUR';
UPDATE employee_master SET area_name='Badami', division_name='Hubli', region_name='Dharwad', area_manager='Basavarajappa', division_manager='Ganapati Fakkirappa Bavukar' WHERE UPPER(branch_name) = 'RAMDURGA';
UPDATE employee_master SET area_name='Ballari', division_name='Hospet', region_name='Chitradurga', area_manager='Bairappa', division_manager='Ambarish M S' WHERE UPPER(branch_name) = 'SANDURU';
UPDATE employee_master SET area_name='Sangareddy', division_name='TS', region_name='TS', area_manager='Sampath Kumar', division_manager='A V Chiranjeevi' WHERE UPPER(branch_name) = 'SANGAREDDY';
UPDATE employee_master SET area_name='Davanagere', division_name='Chitradurga', region_name='Chitradurga', area_manager='Kotresha M', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'SANTHEBENNURU';
UPDATE employee_master SET area_name='Sedam', division_name='Bidar', region_name='Kalburgi', area_manager='Mallikarjun', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'SEDAM';
UPDATE employee_master SET area_name='Shahapur', division_name='Kalburgi', region_name='Kalburgi', area_manager='Basavaraj', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'SHAHAPUR';
UPDATE employee_master SET area_name='Vijayapur', division_name='Kalburgi', region_name='Kalburgi', area_manager='Prakash Biradar', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'SINDAGI';
UPDATE employee_master SET area_name='Lingsugur', division_name='Bidar', region_name='Kalburgi', area_manager='Nagaraj Itagi', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'SINDHNUR';
UPDATE employee_master SET area_name='Tumkur', division_name='Tumkur', region_name='Tumkur', area_manager='Dilip Kumar', division_manager='Mahesh' WHERE UPPER(branch_name) = 'SIRA';
UPDATE employee_master SET area_name='Ballari', division_name='Hospet', region_name='Chitradurga', area_manager='Bairappa', division_manager='Ambarish M S' WHERE UPPER(branch_name) = 'SIRUGUPPA';
UPDATE employee_master SET area_name='Lingsugur', division_name='Bidar', region_name='Kalburgi', area_manager='Nagaraj Itagi', division_manager='Pandarigonda' WHERE UPPER(branch_name) = 'SIRWAR';
UPDATE employee_master SET area_name='Chikballapura', division_name='Doddaballapura', region_name='Tumkur', area_manager='Ravi', division_manager='Muniraju P V' WHERE UPPER(branch_name) = 'SRINIVASPURA';
UPDATE employee_master SET area_name='Vijayapur', division_name='Kalburgi', region_name='Kalburgi', area_manager='Prakash Biradar', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'TALIKOTI';
UPDATE employee_master SET area_name='Mahaboobnagar', division_name='TS', region_name='TS', area_manager='P Maddilety', division_manager='A V Chiranjeevi' WHERE UPPER(branch_name) = 'TANDUR';
UPDATE employee_master SET area_name='Kadur', division_name='Chitradurga', region_name='Chitradurga', area_manager='Kumar R', division_manager='Arun Kumar K H' WHERE UPPER(branch_name) = 'TARIKERE';
UPDATE employee_master SET area_name='Vijayapur', division_name='Kalburgi', region_name='Kalburgi', area_manager='Prakash Biradar', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'TIKOTA';
UPDATE employee_master SET area_name='Tiptur', division_name='Tumkur', region_name='Tumkur', area_manager='P T Venkatesh', division_manager='Mahesh' WHERE UPPER(branch_name) = 'TIPTUR';
UPDATE employee_master SET area_name='Tumkur', division_name='Tumkur', region_name='Tumkur', area_manager='Dilip Kumar', division_manager='Mahesh' WHERE UPPER(branch_name) = 'TUMKUR';
UPDATE employee_master SET area_name='Tiptur', division_name='Tumkur', region_name='Tumkur', area_manager='P T Venkatesh', division_manager='Mahesh' WHERE UPPER(branch_name) = 'TUREVEKERE';
UPDATE employee_master SET area_name='Vijayapur', division_name='Kalburgi', region_name='Kalburgi', area_manager='Prakash Biradar', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'VIJAYAPUR';
UPDATE employee_master SET area_name='Shahapur', division_name='Kalburgi', region_name='Kalburgi', area_manager='Basavaraj', division_manager='Govind Badiger' WHERE UPPER(branch_name) = 'YADGIR';
UPDATE employee_master SET area_name='Belagavi', division_name='Belagavi', region_name='Dharwad', area_manager='Hajiali Tigadi', division_manager='Sunil Kuber Malage' WHERE UPPER(branch_name) = 'YARAGATTI';
UPDATE employee_master SET area_name='Sangareddy', division_name='TS', region_name='TS', area_manager='Sampath Kumar', division_manager='A V Chiranjeevi' WHERE UPPER(branch_name) = 'ZAHEERABAD';

-- Verification (uncomment to run after COMMIT):
-- SELECT UPPER(branch_name), COUNT(DISTINCT area_name||'|'||division_name||'|'||region_name) AS variants
--   FROM employee_master WHERE branch_name IS NOT NULL
--   GROUP BY UPPER(branch_name) HAVING COUNT(DISTINCT area_name||'|'||division_name||'|'||region_name) > 1;
-- -- Expect 0 rows (all branches should be consistent after migration)

COMMIT;
-- Run ROLLBACK; instead of COMMIT; if verification fails