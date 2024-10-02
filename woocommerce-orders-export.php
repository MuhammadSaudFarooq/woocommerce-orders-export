<?php
/*
Plugin Name: WooCommerce Orders Export
Plugin URI: https://github.com/MuhammadSaudFarooq/woocommerce-orders-export
Description: Export WooCommerce orders to CSV.
Author: Muhammad Saud Farooque
Author URI: https://github.com/MuhammadSaudFarooq
Version: 1.0.0
License: MIT
*/

if (!defined('ABSPATH')) {
    exit; // Exit if accessed directly
}

define("PLUGIN_DIR_URL", plugin_dir_url(__FILE__));
define("PLUGIN_DIR_PATH", plugin_dir_path(__FILE__));

require_once __DIR__ . DIRECTORY_SEPARATOR . 'classes' . DIRECTORY_SEPARATOR . 'WooCommerceOrdersExport.php';

$woocommerceOrdersExport  = new WooCommerce_Orders_Export(__FILE__);