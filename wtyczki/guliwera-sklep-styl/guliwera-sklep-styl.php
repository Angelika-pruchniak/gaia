<?php
/**
 * Plugin Name: Guliwera – wygląd sklepu
 * Description: Strony produktów, koszyk, zamówienie i konto klienta w kolorach i układzie strony guliwera4.pl (nagłówek, menu i stopka jak na pozostałych podstronach). Zostaw włączoną.
 * Version: 1.2
 */
if ( ! defined( 'ABSPATH' ) ) { exit; }

define( 'GK_SS_CHROME_PAGE', 10 );
define( 'GK_SS_MAT_PAGE', 74 );

function gk_ss_is_shop_view() {
	if ( ! function_exists( 'is_woocommerce' ) ) { return false; }
	return is_product() || is_cart() || is_checkout() || is_account_page();
}

// Sklep i kategorie -> strona Materiały (tam jest pełna oferta).
add_action( 'template_redirect', function () {
	if ( ! function_exists( 'is_shop' ) ) { return; }
	if ( is_shop() || is_product_category() || is_product_tag() ) {
		wp_safe_redirect( home_url( '/materialy/' ), 301 );
		exit;
	}
} );

add_filter( 'template_include', function ( $t ) {
	return gk_ss_is_shop_view() ? __DIR__ . '/szablon.php' : $t;
}, 99 );

add_filter( 'body_class', function ( $c ) {
	if ( gk_ss_is_shop_view() ) { $c[] = 'gk-sklep-body'; }
	return $c;
} );

/** Wspólne elementy (style, nagłówek, stopka) brane na żywo ze strony głównej. */
function gk_ss_chrome() {
	static $ch = null;
	if ( null !== $ch ) { return $ch; }
	$c  = (string) get_post_field( 'post_content', GK_SS_CHROME_PAGE );
	$ch = array( 'css' => '', 'head' => '', 'foot' => '' );
	if ( preg_match_all( '#<style[^>]*>.*?</style>#s', $c, $m ) ) { $ch['css'] = implode( "\n", $m[0] ); }
	$a = strpos( $c, '<div class="pasek"' );
	if ( false === $a ) { $a = strpos( $c, '<header class="nav"' ); }
	$b = strpos( $c, '</header>' );
	if ( false !== $a && false !== $b ) { $ch['head'] = substr( $c, $a, $b + 9 - $a ); }
	$a = strpos( $c, '<footer class="stopka"' );
	$b = strpos( $c, '</footer>' );
	if ( false !== $a && false !== $b ) { $ch['foot'] = substr( $c, $a, $b + 9 - $a ); }
	$ch['head'] = str_replace( ' class="on"', '', $ch['head'] );
	$ch['head'] = str_replace( 'href="#', 'href="' . home_url( '/' ) . '#', $ch['head'] );
	$ch['foot'] = str_replace( 'href="#"', 'href="https://www.linkedin.com/company/guliwera-4/"', $ch['foot'] );
	return $ch;
}

/** Okładka produktu (.okl) taka sama jak na stronie Materiały. */
function gk_ss_cover( $id ) {
	$c = (string) get_post_field( 'post_content', GK_SS_MAT_PAGE );
	$p = strpos( $c, 'add-to-cart=' . $id . '"' );
	if ( false !== $p ) {
		$s = strrpos( substr( $c, 0, $p ), '<div class="okl' );
		if ( false !== $s ) {
			$d = strpos( $c, '<div class="d">', $s );
			$e = ( false !== $d ) ? strpos( $c, '</div></div>', $d ) : false;
			if ( false !== $e ) { return substr( $c, $s, $e + 12 - $s ); }
		}
	}
	return '<div class="okl"><span class="m">KIERUNEK ZYSK</span><div class="t"><b>' . esc_html( get_the_title( $id ) ) . '</b></div><div class="d"><span>guliwera4.pl</span><span>PDF</span></div></div>';
}

function gk_ss_typ( $id ) {
	$c = (string) get_post_field( 'post_content', GK_SS_MAT_PAGE );
	$p = strpos( $c, 'add-to-cart=' . $id . '"' );
	if ( false === $p ) { return ''; }
	$s = strrpos( substr( $c, 0, $p ), 'class="typ">' );
	if ( false === $s || $p - $s > 1500 ) { return ''; }
	$e = strpos( $c, '<', $s );
	return trim( substr( $c, $s + 12, $e - $s - 12 ) );
}

function gk_ss_css() {
	return '<style id="gk-sklep">
.gk-sklep-body{margin:0;background:var(--krem)}
.gk .gk-prod{padding:clamp(36px,6vw,80px) 0 clamp(40px,5vw,70px);position:relative;overflow:hidden}
.gk .gk-prod::before{content:"";position:absolute;width:520px;height:520px;border-radius:50%;background:var(--roz-jasny);right:-200px;top:-200px}
.gk .gk-prod::after{content:"";position:absolute;width:300px;height:300px;border-radius:50%;background:var(--zolty-jasny);left:-140px;bottom:-160px}
.gk .gk-prod .wrap{position:relative;z-index:2}
.gk .gk-okr{font-size:14.5px;color:var(--atrament-60);margin-bottom:22px}
.gk .gk-okr a{text-decoration:none;color:var(--koral-ciemny);font-weight:700}
.gk .gk-prod-siatka{display:grid;grid-template-columns:.8fr 1.2fr;gap:clamp(30px,6vw,80px);align-items:center}
@media(max-width:860px){.gk .gk-prod-siatka{grid-template-columns:1fr}}
.gk .gk-prod-okl{max-width:380px;width:100%;margin:0 auto;transform:rotate(-3deg);filter:drop-shadow(0 26px 40px rgba(58,35,64,.22))}
.gk .gk-prod-okl .okl{width:100%}
@media(max-width:860px){.gk .gk-prod-okl{max-width:260px}}
.gk .gk-prod h1{font-size:clamp(34px,5vw,58px);margin:8px 0 14px;max-width:18ch}
.gk .gk-typ{display:inline-block;font-size:13px;font-weight:700;letter-spacing:.1em;background:#fff;color:var(--koral-ciemny);padding:6px 14px;border-radius:999px}
.gk .gk-lead{font-size:clamp(17px,1.6vw,19.5px);max-width:52ch;margin-top:16px}
.gk .gk-cena{font-family:"Fraunces",serif;font-weight:700;font-size:clamp(40px,4.6vw,56px);color:var(--koral);line-height:1;margin:22px 0 4px}
.gk .gk-cena del{font-size:.5em;color:var(--atrament-60);margin-left:10px}
.gk .gk-cena-p{font-size:14.5px;color:var(--atrament-60);margin-bottom:22px}
.gk .gk-ctas{display:flex;flex-wrap:wrap;gap:12px}
.gk .gk-plusy{display:flex;flex-wrap:wrap;gap:8px;margin-top:24px;padding:0;list-style:none}
.gk .gk-plusy li{font-size:14px;font-weight:700;background:#fff;border-radius:999px;padding:8px 14px}
.gk .gk-opis{background:#fff}
.gk .gk-opis-siatka{display:grid;grid-template-columns:1.3fr .7fr;gap:clamp(24px,4vw,50px);align-items:start}
@media(max-width:880px){.gk .gk-opis-siatka{grid-template-columns:1fr}}
.gk .gk-opis-tresc h2,.gk .gk-opis-tresc h3{font-size:clamp(22px,2.4vw,28px);margin:30px 0 14px}
.gk .gk-opis-tresc > :first-child{margin-top:0}
.gk .gk-opis-tresc ul{list-style:none;padding:0;margin:0 0 1.2em}
.gk .gk-opis-tresc li{position:relative;padding:8px 0 8px 34px;border-bottom:1px solid rgba(58,35,64,.08)}
.gk .gk-opis-tresc li::before{content:"✓";position:absolute;left:0;top:8px;width:24px;height:24px;border-radius:50%;background:var(--roz-jasny);color:var(--koral-ciemny);font-weight:700;font-size:13px;display:flex;align-items:center;justify-content:center}
.gk .gk-opis-tresc strong{color:var(--atrament)}
.gk .gk-box{background:var(--liliowy-jasny);border-radius:28px;padding:clamp(24px,3vw,34px);position:sticky;top:110px}
.gk .gk-box h3{font-size:24px;margin-bottom:10px}
.gk .gk-box ol{padding-left:20px;margin:0 0 20px}
.gk .gk-box li{margin-bottom:8px}
.gk .gk-box .gk-cena{font-size:38px;margin-top:6px}
.gk .gk-inne .mini{display:grid;grid-template-columns:repeat(3,1fr);gap:18px}
@media(max-width:820px){.gk .gk-inne .mini{grid-template-columns:1fr}}
.gk .gk-inne .mini a{display:grid;grid-template-columns:70px 1fr;gap:16px;align-items:center;background:#fff;border-radius:20px;padding:14px;text-decoration:none;box-shadow:0 10px 26px rgba(58,35,64,.06);transition:transform .15s}
.gk .gk-inne .mini a:hover{transform:translateY(-3px)}
.gk .gk-inne .mini .okl{padding:10%;border-radius:6px}
.gk .gk-inne .mini .okl b{font-size:9px;text-decoration:none}
.gk .gk-inne .mini .okl .m,.gk .gk-inne .mini .okl .d,.gk .gk-inne .mini .okl .tag{display:none}
.gk .gk-inne .mini h4{font-size:17px;margin-bottom:4px}
.gk .gk-inne .mini span.c{font-family:"Fraunces",serif;font-weight:700;color:var(--koral)}
.gk .naglowek.srodek{text-align:center;margin:0 auto clamp(30px,4vw,48px);max-width:640px}
.gk .gk-inne .mini a{color:var(--atrament)}
.gk .gk-wc{padding:clamp(36px,5vw,70px) 0 clamp(56px,7vw,100px)}
.gk .gk-wc h1{font-size:clamp(34px,5vw,56px);margin:6px 0 28px}
.gk .gk-wc-karta{background:#fff;border-radius:28px;padding:clamp(20px,3.4vw,44px);box-shadow:0 16px 44px rgba(58,35,64,.08)}
.gk .gk-wc input[type=text],.gk .gk-wc input[type=email],.gk .gk-wc input[type=tel],.gk .gk-wc input[type=password],.gk .gk-wc input[type=number],.gk .gk-wc select,.gk .gk-wc textarea,.gk .gk-wc .wc-block-components-text-input input,.gk .gk-wc .wc-block-components-combobox .wc-block-components-combobox-control input{border:2px solid var(--brzoskwinia) !important;border-radius:14px !important;background:var(--krem) !important;font-family:"DM Sans",sans-serif !important;color:var(--atrament) !important}
.gk .gk-wc input:focus,.gk .gk-wc select:focus,.gk .gk-wc textarea:focus{border-color:var(--koral) !important;outline:none;box-shadow:0 0 0 3px rgba(255,107,61,.15) !important}
.gk .gk-wc .button,.gk .gk-wc button[type=submit],.gk .gk-wc .wc-block-components-button:not(.is-link),.gk .gk-wc .wc-block-cart__submit-button,.gk .gk-wc .wc-block-components-checkout-place-order-button,.gk .gk-wc #place_order,.gk .gk-wc .checkout-button{background:var(--koral) !important;border:2.5px solid var(--koral) !important;color:#fff !important;border-radius:999px !important;font-family:"DM Sans",sans-serif !important;font-weight:700 !important;font-size:16px !important;padding:14px 28px !important;text-decoration:none !important;box-shadow:none !important;text-transform:none !important;letter-spacing:0 !important;min-height:0 !important}
.gk .gk-wc .button:hover,.gk .gk-wc button[type=submit]:hover,.gk .gk-wc .wc-block-components-button:not(.is-link):hover,.gk .gk-wc #place_order:hover{background:var(--koral-ciemny) !important;border-color:var(--koral-ciemny) !important}
.gk .gk-wc a{color:var(--koral-ciemny)}
.gk .gk-wc table{width:100%;border-collapse:collapse}
.gk .gk-wc th,.gk .gk-wc td{padding:12px 10px;border-bottom:1px solid rgba(58,35,64,.1);text-align:left}
.gk .gk-wc .wc-block-components-totals-item__value,.gk .gk-wc .wc-block-formatted-money-amount,.gk .gk-wc .woocommerce-Price-amount{font-weight:700}
.gk .gk-wc .wc-block-components-order-summary,.gk .gk-wc .wc-block-components-sidebar{border-radius:22px}
.gk .gk-wc .wc-block-components-sidebar-layout .wc-block-components-sidebar{background:var(--krem);padding:20px;border-radius:22px}
.gk .gk-wc .wc-block-components-title,.gk .gk-wc h2,.gk .gk-wc h3{font-family:"Fraunces",serif !important;color:var(--atrament) !important;text-transform:none !important}
.gk .gk-wc .wc-block-components-radio-control__option,.gk .gk-wc .wc_payment_methods li{border-radius:14px}
.gk .gk-wc .woocommerce-info,.gk .gk-wc .woocommerce-message,.gk .gk-wc .woocommerce-error,.gk .gk-wc .wc-block-components-notice-banner{border-radius:16px !important;border:0 !important;background:var(--zolty-jasny) !important;color:var(--atrament) !important;padding:14px 18px !important;list-style:none}
.gk .gk-wc .wc-block-cart__empty-cart__title,.gk .gk-wc .cart-empty{font-family:"Fraunces",serif;font-size:26px;text-align:center}
.gk .gk-wc .wp-block-woocommerce-empty-cart-block .wp-block-separator,.gk .gk-wc .wp-block-woocommerce-empty-cart-block .wp-block-heading:not(.wc-block-cart__empty-cart__title),.gk .gk-wc .wp-block-woocommerce-empty-cart-block .wp-block-woocommerce-product-new{display:none}
.gk .gk-pusty{text-align:center;margin-top:24px}
.gk .gk-edytuj{margin:-14px 0 20px;font-size:16px}
.gk .gk-edytuj a{color:var(--koral-ciemny);font-weight:700;text-decoration:none}
.gk .gk-wc .woocommerce-MyAccount-navigation ul{list-style:none;padding:0;display:flex;flex-wrap:wrap;gap:8px;margin-bottom:24px}
.gk .gk-wc .woocommerce-MyAccount-navigation a{display:block;padding:8px 16px;border-radius:999px;background:var(--krem);text-decoration:none;font-weight:700;color:var(--atrament)}
.gk .gk-wc .woocommerce-MyAccount-navigation .is-active a{background:var(--koral);color:#fff}
</style>';
}

register_activation_hook( __FILE__, function () {
	if ( function_exists( 'w3tc_flush_all' ) ) { w3tc_flush_all(); }
} );

/** Ustawienia sklepu dla Polski (raz, po aktualizacji wtyczki). */
add_action( 'init', function () {
	if ( '1.1' === get_option( 'gk_ss_ver' ) || ! function_exists( 'WC' ) ) { return; }
	update_option( 'woocommerce_default_country', 'PL' );
	update_option( 'woocommerce_default_customer_address', 'base' );
	update_option( 'woocommerce_currency', 'PLN' );
	update_option( 'woocommerce_currency_pos', 'right_space' );
	update_option( 'woocommerce_price_thousand_sep', ' ' );
	update_option( 'woocommerce_price_decimal_sep', ',' );
	update_option( 'woocommerce_price_num_decimals', '2' );
	update_option( 'woocommerce_checkout_address_2_field', 'hidden' );
	update_option( 'woocommerce_checkout_phone_field', 'hidden' );
	update_option( 'woocommerce_checkout_company_field', 'optional' );
	update_option( 'gk_ss_ver', '1.1' );
	if ( function_exists( 'w3tc_flush_all' ) ) { w3tc_flush_all(); }
} );

/** Same pliki do pobrania: bez adresu w zamówieniu (imię, nazwisko, e-mail, kraj zostają). */
function gk_ss_hide_address( $loc ) {
	foreach ( array( 'address_1', 'address_2', 'city', 'state', 'postcode' ) as $k ) {
		$loc[ $k ] = array_merge( isset( $loc[ $k ] ) ? (array) $loc[ $k ] : array(), array( 'required' => false, 'hidden' => true ) );
	}
	return $loc;
}
add_filter( 'woocommerce_get_country_locale_default', 'gk_ss_hide_address', 99 );
add_filter( 'woocommerce_get_country_locale_base', 'gk_ss_hide_address', 99 );
add_filter( 'woocommerce_get_country_locale', function ( $all ) {
	foreach ( $all as $c => $loc ) { $all[ $c ] = gk_ss_hide_address( $loc ); }
	foreach ( array( 'PL', 'US', 'GB', 'DE' ) as $c ) { if ( ! isset( $all[ $c ] ) ) { $all[ $c ] = gk_ss_hide_address( array() ); } }
	return $all;
}, 99 );
add_filter( 'woocommerce_billing_fields', function ( $f ) {
	foreach ( array( 'billing_address_1', 'billing_address_2', 'billing_city', 'billing_state', 'billing_postcode' ) as $k ) { unset( $f[ $k ] ); }
	return $f;
}, 99 );

/** Każdy PDF tylko raz w koszyku: ustaw „sprzedawany pojedynczo” (raz, wersja 1.2). */
add_action( 'init', function () {
	if ( '1.2' === get_option( 'gk_ss_ver2' ) || ! function_exists( 'wc_get_product' ) ) { return; }
	foreach ( array( 67, 68, 69, 70, 71, 72, 73 ) as $id ) {
		$p = wc_get_product( $id );
		if ( $p ) { $p->set_sold_individually( true ); $p->save(); }
	}
	update_option( 'gk_ss_ver2', '1.2' );
	if ( function_exists( 'w3tc_flush_all' ) ) { w3tc_flush_all(); }
}, 20 );

/** Produkt już w koszyku: nie dodawaj drugi raz i nie pokazuj błędu. */
add_filter( 'woocommerce_add_to_cart_validation', function ( $ok, $product_id ) {
	if ( ! $ok || ! function_exists( 'WC' ) || ! WC()->cart ) { return $ok; }
	foreach ( WC()->cart->get_cart() as $item ) {
		if ( (int) $item['product_id'] === (int) $product_id ) { return false; }
	}
	return $ok;
}, 1, 2 );

/** Wyczyść ewentualne komunikaty „nie można dodać kolejnego” dla produktów pojedynczych. */
add_filter( 'woocommerce_cart_product_cannot_add_another_message', '__return_empty_string' );

/** Po dodaniu z linku ?add-to-cart przekieruj na czysty adres, żeby odświeżenie nic nie dodawało. */
add_action( 'woocommerce_add_to_cart', function () {
	if ( isset( $_GET['add-to-cart'] ) && ! wp_doing_ajax() && function_exists( 'wc_get_checkout_url' ) ) {
		add_filter( 'woocommerce_add_to_cart_redirect', function () { return wc_get_checkout_url(); } );
	}
} );

/** Napraw koszyki, w których PDF jest już 2 razy lub więcej: zostaw 1 sztukę. */
add_action( 'woocommerce_cart_loaded_from_session', function ( $cart ) {
	foreach ( $cart->get_cart() as $key => $item ) {
		if ( $item['quantity'] > 1 ) { $cart->set_quantity( $key, 1, false ); }
	}
}, 20 );
