<?php

require_once PLUGIN_DIR_PATH . 'spreadsheet/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class WooCommerce_Orders_Export
{
    private $exclude_donation_product_id = null;
    private $exclude_bulk_product_id = null;
    private $exclude_single_product_id = null;

    public function __construct()
    {
        // Hook for logged-in users
        add_action("wp_ajax_orders_csv", [$this, 'orders_csv_fn']);
        // Hook for guest users
        add_action("wp_ajax_nopriv_orders_csv", [$this, 'orders_csv_fn']);

        // Enqueue scripts
        // add_action('wp_enqueue_scripts', [$this, 'enqueue_scripts']);

        $this->exclude_donation_product_id = 893;
        $this->exclude_bulk_product_id = 27;
        $this->exclude_single_product_id = 14;
    }

    // Function to handle AJAX request
    public function orders_csv_fn()
    {
        if (isset($_GET['action']) && $_GET['action'] == 'orders_csv') {
            // Start output buffering to avoid any unexpected output
            ob_start();

            // Create a new spreadsheet
            $spreadsheet = new Spreadsheet();

            // Add first sheet (Single Orders)
            $sheet1 = $spreadsheet->getActiveSheet();
            $sheet1->setTitle('Single Orders');
            $sheet1->fromArray([
                'First Name',
                'Last Name',
                'Email',
                'Amount',
                'Address',
                'Address 2',
                'City',
                'State',
                'Zip Code',
                'Country Code',
                'Country Dialing Code',
                'Reference Details',
                'Phone Number',
                'Number of Cards',
                'Amount per Card'
            ], NULL, 'A1');

            // Get WooCommerce orders
            /* $args1 = [
                'limit' => -1,
            ]; */
            $args1 = [
                'limit'        => -1,  // No limit on the number of orders
                'date_created' => '>=' . (new WC_DateTime())->modify('-1 day')->date('Y-m-d H:i:s'),  // Orders from the last 24 hours
            ];
            $orders1 = wc_get_orders($args1);

            if (empty($orders1)) {
                error_log('No orders found. Check the wc_get_orders query.');
            } else {
                $row = 2; // Start at the second row
                foreach ($orders1 as $order) {
                    // Check if the order contains only one product
                    $items = $order->get_items();
                    if (count($items) != 1) {
                        continue; // Skip if the order has more than one product
                    }

                    // Check if the order has a subscription (skip if it does)
                    /* $isMember = wcs_get_subscriptions_for_order($order->get_id());
                    if (!empty($isMember)) {
                        continue; // Skip if this order is linked to a subscription
                    } */

                    // Loop through the items to check for the excluded product
                    $skip_order = false;
                    foreach ($items as $item) {
                        $product_id = $item->get_product_id();

                        // Skip order if it contains the excluded donation product
                        if ($product_id == $this->exclude_donation_product_id) {
                            $skip_order = true;
                            break;
                        }

                        // Skip order if it contains the excluded bulk product
                        if ($product_id == $this->exclude_bulk_product_id) {
                            $skip_order = true;
                            break;
                        }

                        if (empty($isMember)) {
                            $skip_order = true; // Set flag to skip this order
                            break;
                        }
                    }

                    if ($skip_order) {
                        continue; // Skip this order if it matches any exclusion criteria
                    }

                    // Extract order details
                    $first_name = $order->get_billing_first_name();
                    $last_name = $order->get_billing_last_name();
                    $email = $order->get_billing_email();
                    $amount = $order->get_total();
                    $address_1 = $order->get_billing_address_1();
                    $address_2 = $order->get_billing_address_2();
                    $city = $order->get_billing_city();
                    $state = $order->get_billing_state();
                    $zip_code = $order->get_billing_postcode();
                    $country_code = $order->get_billing_country();
                    $phone = $order->get_billing_phone();
                    $reference_details = $order->get_id();
                    $state_fullname = WC()->countries->get_states($country_code)[$state];

                    // Number of cards and amount per card (custom logic)
                    $number_of_cards = 1; // Assuming it's one card per single-product order
                    $amount_per_card = $amount / $number_of_cards;

                    // Country dialing code (custom function or static mapping)
                    $dialing_code = ($country_code != '') ? $this->get_country_dialing_code($country_code) : $this->get_country_dialing_code($state);

                    // Add row to spreadsheet
                    $sheet1->fromArray([
                        $first_name,
                        $last_name,
                        $email,
                        $amount,
                        $address_1,
                        $address_2,
                        $city,
                        $state_fullname,
                        $zip_code,
                        $country_code,
                        $dialing_code,
                        $reference_details,
                        $phone,
                        $number_of_cards,
                        $amount_per_card
                    ], NULL, 'A' . $row);

                    $row++;
                }
            }

            // Add second sheet (Bulk Orders)
            $sheet2 = $spreadsheet->createSheet();
            $sheet2->setTitle('Bulk Orders');
            $sheet2->fromArray([
                'First Name',
                'Last Name',
                'Email',
                'Amount',
                'Address',
                'Address 2',
                'City',
                'State',
                'Zip Code',
                'Country Code',
                'Country Dialing Code',
                'Reference Detals',
                'Phone Number',
                'Number of Cards',
                'Amount per Card'
            ], NULL, 'A1');

            // Get WooCommerce orders
            /* $args2 = [
                'limit' => -1,
            ]; */
            $args2 = [
                'limit'        => -1,  // No limit on the number of orders
                'date_created' => '>=' . (new WC_DateTime())->modify('-1 day')->date('Y-m-d H:i:s'),  // Orders from the last 24 hours
            ];
            $orders2 = wc_get_orders($args2);

            if (empty($orders2)) {
                error_log('No orders found. Check the wc_get_orders query.');
            } else {
                $row = 2; // Start at the second row
                foreach ($orders2 as $order) {
                    // Check if the order contains only one product
                    $items = $order->get_items();
                    if (count($items) != 1) {
                        continue; // Skip if the order has more than one product
                    }

                    // Check membership
                    $isMember = wcs_get_subscriptions_for_order($order->get_id());

                    // Loop through the items to check for the excluded product
                    $skip_order = false;
                    $bulk_count = 0;
                    foreach ($items as $item) {
                        if ($item->get_product_id() == $this->exclude_donation_product_id) {
                            $skip_order = true; // Set flag to skip this order
                            break;
                        }

                        if ($item->get_product_id() == $this->exclude_single_product_id) {
                            $skip_order = true; // Set flag to skip this order
                            break;
                        }

                        if (!empty($isMember)) {
                            $skip_order = true; // Set flag to skip this order
                            break;
                        }

                        $product_data = $item->get_product();
                        $bulk_count = $product_data->get_attributes()['pa_package'];
                    }
                    if ($skip_order) {
                        continue; // Skip this order if it contains the excluded product
                    }

                    // Extract order details
                    $first_name = $order->get_billing_first_name();
                    $last_name = $order->get_billing_last_name();
                    $email = $order->get_billing_email();
                    $amount = number_format($order->get_total() * (int)$bulk_count, 2);
                    $address_1 = $order->get_billing_address_1();
                    $address_2 = $order->get_billing_address_2();
                    $city = $order->get_billing_city();
                    $state = $order->get_billing_state();
                    $zip_code = $order->get_billing_postcode();
                    $country_code = $order->get_billing_country();
                    $phone = $order->get_billing_phone();
                    $reference_details = $order->get_id();
                    $state_fullname = WC()->countries->get_states($country_code)[$state];

                    // Number of cards and amount per card (custom logic)
                    $number_of_cards = $bulk_count; // Assuming it's one card per single-product order
                    $amount_per_card = $order->get_total();

                    // Country dialing code (custom function or static mapping)
                    $dialing_code = ($country_code != '') ? $this->get_country_dialing_code($country_code) : $this->get_country_dialing_code($state); // Define or hard-code this function
                    // $dialing_code = '123'; // Define or hard-code this function

                    // Add row to spreadsheet
                    $sheet2->fromArray([
                        $first_name,
                        $last_name,
                        $email,
                        $amount,
                        $address_1,
                        $address_2,
                        $city,
                        $state_fullname,
                        $zip_code,
                        $country_code,
                        $dialing_code,
                        $reference_details,
                        $phone,
                        $number_of_cards,
                        $amount_per_card
                    ], NULL, 'A' . $row);

                    $row++;
                }
            }

            // Add third sheet (Membership Orders)
            $sheet3 = $spreadsheet->createSheet();
            $sheet3->setTitle('Membership Orders');
            $sheet3->fromArray([
                'First Name',
                'Last Name',
                'Email',
                'Amount',
                'Address',
                'Address 2',
                'City',
                'State',
                'Zip Code',
                'Country Code',
                'Country Dialing Code',
                'Reference Detals',
                'Phone Number',
                'Number of Cards',
                'Amount per Card'
            ], NULL, 'A1');

            // Get WooCommerce orders
            /* $args3 = [
                'limit' => -1,
            ]; */
            $args3 = [
                'limit'        => -1,  // No limit on the number of orders
                'date_created' => '>=' . (new WC_DateTime())->modify('-1 day')->date('Y-m-d H:i:s'),  // Orders from the last 24 hours
            ];
            $orders3 = wc_get_orders($args3);

            if (empty($orders3)) {
                error_log('No orders found. Check the wc_get_orders query.');
            } else {
                $row = 2; // Start at the second row
                foreach ($orders3 as $order) {
                    // Check if the order contains only one product
                    $items = $order->get_items();
                    if (count($items) != 1) {
                        continue; // Skip if the order has more than one product
                    }

                    // Check membership
                    $isMember = wcs_get_subscriptions_for_order($order->get_id());

                    // Loop through the items to check for the excluded product
                    $skip_order = false;
                    foreach ($items as $item) {
                        if ($item->get_product_id() == $this->exclude_donation_product_id) {
                            $skip_order = true; // Set flag to skip this order
                            break;
                        }

                        if (empty($isMember)) {
                            $skip_order = true; // Set flag to skip this order
                            break;
                        }
                    }
                    if ($skip_order) {
                        continue; // Skip this order if it contains the excluded product
                    }

                    // Extract order details
                    $first_name = $order->get_billing_first_name();
                    $last_name = $order->get_billing_last_name();
                    $email = $order->get_billing_email();
                    $amount = $order->get_total();
                    $address_1 = $order->get_billing_address_1();
                    $address_2 = $order->get_billing_address_2();
                    $city = $order->get_billing_city();
                    $state = $order->get_billing_state();
                    $zip_code = $order->get_billing_postcode();
                    $country_code = $order->get_billing_country();
                    $phone = $order->get_billing_phone();
                    $reference_details = $order->get_id();
                    $state_fullname = WC()->countries->get_states($country_code)[$state];

                    // Number of cards and amount per card (custom logic)
                    $number_of_cards = 1; // Assuming it's one card per single-product order
                    $amount_per_card = $amount / $number_of_cards;

                    // Country dialing code (custom function or static mapping)
                    $dialing_code = ($country_code != '') ? $this->get_country_dialing_code($country_code) : $this->get_country_dialing_code($state); // Define or hard-code this function
                    // $dialing_code = '123'; // Define or hard-code this function

                    // Add row to spreadsheet
                    $sheet3->fromArray([
                        $first_name,
                        $last_name,
                        $email,
                        $amount,
                        $address_1,
                        $address_2,
                        $city,
                        $state_fullname,
                        $zip_code,
                        $country_code,
                        $dialing_code,
                        $reference_details,
                        $phone,
                        $number_of_cards,
                        $amount_per_card
                    ], NULL, 'A' . $row);

                    $row++;
                }
            }

            // Set headers to force download of XLSX
            /* header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment; filename="orders-export-' . date('Y-m-d-H-i-s') . '.xlsx";');
            header('Cache-Control: max-age=0');

            // Write the spreadsheet to output
            $writer = new Xlsx($spreadsheet);
            $writer->save('php://output');

            // Clean output buffer and flush content
            ob_flush(); // Ensure all output is sent
            exit; */


            // Save spreadsheet to a temporary file
            $upload_dir = wp_upload_dir();
            $file_path = $upload_dir['basedir'] . '/orders-export-' . date('Y-m-d-H-i-s') . '.xlsx';
            $writer = new Xlsx($spreadsheet);
            $writer->save($file_path);

            // Email setup
            $email_to = 'muhammad.saud@koderlabs.com'; // Change to recipient email address
            $subject = 'Exported Orders CSV';
            $message = 'Exported Orders CSV.';
            $headers = ['Content-Type: text/html; charset=UTF-8'];

            if (isset($_GET['emailAddr']) && $_GET['emailAddr'] != '') {
                $email_to = $_GET['emailAddr'];
            }

            // Attach the spreadsheet file
            $attachments = [$file_path];

            // Send the email with the attachment
            $mail_sent = wp_mail($email_to, $subject, $message, $headers, $attachments);

            if ($mail_sent) {
                // Log success or handle any post-email actions
                error_log('Email sent successfully with the attachment.');
            } else {
                // Log failure or handle errors
                error_log('Failed to send email.');
            }

            // Clean up - delete the temporary file after sending
            if (file_exists($file_path)) {
                unlink($file_path);
            }

            // Clean output buffer and flush content
            ob_flush(); // Ensure all output is sent
            exit;
        }

        exit; // Terminate to prevent additional output
    }

    public function get_country_dialing_code($country_code)
    {
        $country_dialing_codes = [
            'AF' => '+93',    // Afghanistan
            'AL' => '+355',   // Albania
            'DZ' => '+213',   // Algeria
            'AS' => '+1-684', // American Samoa
            'AD' => '+376',   // Andorra
            'AO' => '+244',   // Angola
            'AR' => '+54',    // Argentina
            'AM' => '+374',   // Armenia
            'GA' => '+445',   // Atlanta
            'AU' => '+61',    // Australia
            'AT' => '+43',    // Austria
            'AZ' => '+994',   // Azerbaijan
            'BH' => '+973',   // Bahrain
            'BD' => '+880',   // Bangladesh
            'BY' => '+375',   // Belarus
            'BE' => '+32',    // Belgium
            'BZ' => '+501',   // Belize
            'BJ' => '+229',   // Benin
            'BT' => '+975',   // Bhutan
            'BO' => '+591',   // Bolivia
            'BA' => '+387',   // Bosnia and Herzegovina
            'BW' => '+267',   // Botswana
            'BR' => '+55',    // Brazil
            'BN' => '+673',   // Brunei
            'BG' => '+359',   // Bulgaria
            'BF' => '+226',   // Burkina Faso
            'BI' => '+257',   // Burundi
            'KH' => '+855',   // Cambodia
            'CM' => '+237',   // Cameroon
            'CA' => '+1',     // Canada
            'CV' => '+238',   // Cape Verde
            'CF' => '+236',   // Central African Republic
            'TD' => '+235',   // Chad
            'CL' => '+56',    // Chile
            'CN' => '+86',    // China
            'CO' => '+57',    // Colombia
            'KM' => '+269',   // Comoros
            'CG' => '+242',   // Congo
            'CR' => '+506',   // Costa Rica
            'HR' => '+385',   // Croatia
            'CU' => '+53',    // Cuba
            'CY' => '+357',   // Cyprus
            'CZ' => '+420',   // Czech Republic
            'DK' => '+45',    // Denmark
            'DJ' => '+253',   // Djibouti
            'DM' => '+1-767', // Dominica
            'DO' => '+1-809', // Dominican Republic
            'EC' => '+593',   // Ecuador
            'EG' => '+20',    // Egypt
            'SV' => '+503',   // El Salvador
            'GQ' => '+240',   // Equatorial Guinea
            'ER' => '+291',   // Eritrea
            'EE' => '+372',   // Estonia
            'ET' => '+251',   // Ethiopia
            'FJ' => '+679',   // Fiji
            'FI' => '+358',   // Finland
            'FR' => '+33',    // France
            'GA' => '+241',   // Gabon
            'GM' => '+220',   // Gambia
            'GE' => '+995',   // Georgia
            'DE' => '+49',    // Germany
            'GH' => '+233',   // Ghana
            'GR' => '+30',    // Greece
            'GD' => '+1-473', // Grenada
            'GT' => '+502',   // Guatemala
            'GN' => '+224',   // Guinea
            'GW' => '+245',   // Guinea-Bissau
            'GY' => '+592',   // Guyana
            'HT' => '+509',   // Haiti
            'HN' => '+504',   // Honduras
            'HU' => '+36',    // Hungary
            'IS' => '+354',   // Iceland
            'IN' => '+91',    // India
            'ID' => '+62',    // Indonesia
            'IR' => '+98',    // Iran
            'IQ' => '+964',   // Iraq
            'IE' => '+353',   // Ireland
            'IL' => '+972',   // Israel
            'IT' => '+39',    // Italy
            'JM' => '+1-876', // Jamaica
            'JP' => '+81',    // Japan
            'JO' => '+962',   // Jordan
            'KZ' => '+7',     // Kazakhstan
            'KE' => '+254',   // Kenya
            'KI' => '+686',   // Kiribati
            'KP' => '+850',   // North Korea
            'KR' => '+82',    // South Korea
            'KW' => '+965',   // Kuwait
            'KG' => '+996',   // Kyrgyzstan
            'LA' => '+856',   // Laos
            'LV' => '+371',   // Latvia
            'LB' => '+961',   // Lebanon
            'LS' => '+266',   // Lesotho
            'LR' => '+231',   // Liberia
            'LY' => '+218',   // Libya
            'LI' => '+423',   // Liechtenstein
            'LT' => '+370',   // Lithuania
            'LU' => '+352',   // Luxembourg
            'MG' => '+261',   // Madagascar
            'MW' => '+265',   // Malawi
            'MY' => '+60',    // Malaysia
            'MV' => '+960',   // Maldives
            'ML' => '+223',   // Mali
            'MT' => '+356',   // Malta
            'MH' => '+692',   // Marshall Islands
            'MR' => '+222',   // Mauritania
            'MU' => '+230',   // Mauritius
            'MX' => '+52',    // Mexico
            'FM' => '+691',   // Micronesia
            'MD' => '+373',   // Moldova
            'MC' => '+377',   // Monaco
            'MN' => '+976',   // Mongolia
            'ME' => '+382',   // Montenegro
            'MA' => '+212',   // Morocco
            'MZ' => '+258',   // Mozambique
            'MM' => '+95',    // Myanmar
            'NA' => '+264',   // Namibia
            'NR' => '+674',   // Nauru
            'NP' => '+977',   // Nepal
            'NL' => '+31',    // Netherlands
            'NZ' => '+64',    // New Zealand
            'NI' => '+505',   // Nicaragua
            'NE' => '+227',   // Niger
            'NG' => '+234',   // Nigeria
            'NO' => '+47',    // Norway
            'OM' => '+968',   // Oman
            'PK' => '+92',    // Pakistan
            'PW' => '+680',   // Palau
            'PA' => '+507',   // Panama
            'PG' => '+675',   // Papua New Guinea
            'PY' => '+595',   // Paraguay
            'PE' => '+51',    // Peru
            'PH' => '+63',    // Philippines
            'PL' => '+48',    // Poland
            'PT' => '+351',   // Portugal
            'QA' => '+974',   // Qatar
            'RO' => '+40',    // Romania
            'RU' => '+7',     // Russia
            'RW' => '+250',   // Rwanda
            'KN' => '+1-869', // Saint Kitts and Nevis
            'LC' => '+1-758', // Saint Lucia
            'VC' => '+1-784', // Saint Vincent and the Grenadines
            'WS' => '+685',   // Samoa
            'SM' => '+378',   // San Marino
            'ST' => '+239',   // Sao Tome and Principe
            'SA' => '+966',   // Saudi Arabia
            'SN' => '+221',   // Senegal
            'RS' => '+381',   // Serbia
            'SC' => '+248',   // Seychelles
            'SL' => '+232',   // Sierra Leone
            'SG' => '+65',    // Singapore
            'SK' => '+421',   // Slovakia
            'SI' => '+386',   // Slovenia
            'SB' => '+677',   // Solomon Islands
            'SO' => '+252',   // Somalia
            'ZA' => '+27',    // South Africa
            'ES' => '+34',    // Spain
            'LK' => '+94',    // Sri Lanka
            'SD' => '+249',   // Sudan
            'SR' => '+597',   // Suriname
            'SZ' => '+268',   // Swaziland
            'SE' => '+46',    // Sweden
            'CH' => '+41',    // Switzerland
            'SY' => '+963',   // Syria
            'TW' => '+886',   // Taiwan
            'TJ' => '+992',   // Tajikistan
            'TZ' => '+255',   // Tanzania
            'TH' => '+66',    // Thailand
            'TG' => '+228',   // Togo
            'TO' => '+676',   // Tonga
            'TT' => '+1-868', // Trinidad and Tobago
            'TN' => '+216',   // Tunisia
            'TR' => '+90',    // Turkey
            'TM' => '+993',   // Turkmenistan
            'UG' => '+256',   // Uganda
            'UA' => '+380',   // Ukraine
            'AE' => '+971',   // United Arab Emirates
            'GB' => '+44',    // United Kingdom
            'US' => '+1',     // United States
            'UY' => '+598',   // Uruguay
            'UZ' => '+998',   // Uzbekistan
            'VU' => '+678',   // Vanuatu
            'VE' => '+58',    // Venezuela
            'VN' => '+84',    // Vietnam
            'YE' => '+967',   // Yemen
            'ZM' => '+260',   // Zambia
            'ZW' => '+263',   // Zimbabwe
        ];

        return isset($country_dialing_codes[$country_code]) ? $country_dialing_codes[$country_code] : 'Unknown';
    }
}
