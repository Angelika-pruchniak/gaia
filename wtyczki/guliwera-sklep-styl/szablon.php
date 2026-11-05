<?php
if ( ! defined( 'ABSPATH' ) ) { exit; }
$ch = gk_ss_chrome();
?><!DOCTYPE html>
<html <?php language_attributes(); ?>>
<head>
<meta charset="<?php bloginfo( 'charset' ); ?>">
<meta name="viewport" content="width=device-width, initial-scale=1">
<?php wp_head(); ?>
<?php echo $ch['css']; // phpcs:ignore ?>
<?php echo gk_ss_css(); // phpcs:ignore ?>
</head>
<body <?php body_class(); ?>>
<?php wp_body_open(); ?>
<div class="gk">
<?php
$head = $ch['head'];
if ( is_product() ) {
	$head = str_replace( 'href="' . home_url( '/materialy/' ) . '"', 'class="on" href="' . home_url( '/materialy/' ) . '"', $head );
}
echo $head; // phpcs:ignore

if ( is_product() ) :
	while ( have_posts() ) :
		the_post();
		$id = get_the_ID();
		$pr = wc_get_product( $id );
		$kup = home_url( '/zamowienie/?add-to-cart=' . $id );
		$typ = gk_ss_typ( $id );
		$is_pak = ( 67 === $id );
		?>
<section class="gk-prod"><div class="wrap">
  <p class="gk-okr"><a href="<?php echo esc_url( home_url( '/materialy/' ) ); ?>">← Wszystkie materiały</a> · Kierunek Zysk</p>
  <?php if ( function_exists( 'wc_print_notices' ) ) { wc_print_notices(); } ?>
  <div class="gk-prod-siatka">
    <div class="gk-prod-okl"><?php echo gk_ss_cover( $id ); // phpcs:ignore ?></div>
    <div>
      <span class="pisane"><?php echo $is_pak ? 'najczęściej wybierane' : 'materiał do pobrania'; ?></span>
      <h1><?php the_title(); ?></h1>
      <?php if ( $typ || $is_pak ) : ?><span class="gk-typ"><?php echo esc_html( $is_pak ? 'PLIK ZIP · 7 MATERIAŁÓW' : $typ ); ?></span><?php endif; ?>
      <?php if ( $pr && $pr->get_short_description() ) : ?><p class="gk-lead"><?php echo wp_kses_post( $pr->get_short_description() ); ?></p><?php endif; ?>
      <div class="gk-cena"><?php echo $pr ? wp_kses_post( wc_price( $pr->get_price(), array( 'decimals' => 0 ) ) ) : ''; ?><?php if ( $is_pak ) : ?><del>432 zł</del><?php endif; ?></div>
      <p class="gk-cena-p">jednorazowo, cena brutto · pobierasz od razu po płatności</p>
      <div class="gk-ctas">
        <a class="btn" href="<?php echo esc_url( $kup ); ?>"><?php echo $is_pak ? 'Kup pakiet i pobierz' : 'Kup i pobierz'; ?></a>
        <?php if ( ! $is_pak ) : ?><a class="btn btn-obrys" href="<?php echo esc_url( get_permalink( 67 ) ); ?>">Cały pakiet za 297 zł</a><?php else : ?><a class="btn btn-obrys" href="<?php echo esc_url( home_url( '/materialy/#sklep' ) ); ?>">Pojedyncze PDF-y</a><?php endif; ?>
      </div>
      <ul class="gk-plusy"><li>✦ Plik od razu na maila</li><li>✦ 5 pobrań przez 30 dni</li><li>✦ Bez zakładania konta</li></ul>
    </div>
  </div>
</div></section>

<section class="sekcja gk-opis"><div class="wrap"><div class="gk-opis-siatka">
  <div class="gk-opis-tresc">
    <span class="pisane">co jest w środku?</span>
    <h2 style="margin-top:6px">Opis materiału</h2>
    <?php the_content(); ?>
  </div>
  <aside class="gk-box">
    <h3>Jak to działa?</h3>
    <ol>
      <li>Klikasz „Kup i pobierz”.</li>
      <li>Podajesz imię i e-mail, płacisz.</li>
      <li>Link do pobrania dostajesz od razu na ekranie i mailem.</li>
    </ol>
    <div class="gk-cena"><?php echo $pr ? wp_kses_post( wc_price( $pr->get_price(), array( 'decimals' => 0 ) ) ) : ''; ?></div>
    <a class="btn" style="margin-top:14px" href="<?php echo esc_url( $kup ); ?>">Kup i pobierz</a>
  </aside>
</div></div></section>

<section class="sekcja gk-inne"><div class="wrap">
  <div class="naglowek srodek"><span class="pisane">może przyda się też</span><h2>Inne materiały Kierunek Zysk</h2></div>
  <div class="mini">
		<?php
		$n = 0;
		foreach ( array( 67, 68, 69, 70, 71, 72, 73 ) as $oid ) {
			if ( $oid === $id || $n >= 3 ) { continue; }
			$op = wc_get_product( $oid );
			if ( ! $op || 'publish' !== $op->get_status() ) { continue; }
			$n++;
			echo '<a href="' . esc_url( get_permalink( $oid ) ) . '">' . gk_ss_cover( $oid ) . '<div><h4>' . esc_html( $op->get_name() ) . '</h4><span class="c">' . wp_kses_post( wc_price( $op->get_price(), array( 'decimals' => 0 ) ) ) . '</span></div></a>'; // phpcs:ignore
		}
		?>
  </div>
</div></section>
		<?php
	endwhile;
else :
	while ( have_posts() ) :
		the_post();
		$pis = is_cart() ? 'Twój koszyk' : ( is_account_page() ? 'Twoje konto' : 'prawie gotowe' );
		if ( is_checkout() && function_exists( 'is_order_received_page' ) && is_order_received_page() ) { $pis = 'dziękuję!'; }
		?>
<section class="gk-wc"><div class="wrap">
  <span class="pisane"><?php echo esc_html( $pis ); ?></span>
  <h1><?php the_title(); ?></h1>
  <?php if ( is_checkout() && ! ( function_exists( 'is_order_received_page' ) && is_order_received_page() ) ) : ?>
  <p class="gk-edytuj"><a href="<?php echo esc_url( wc_get_cart_url() ); ?>">← Zmień lub usuń produkty w koszyku</a></p>
  <?php endif; ?>
  <div class="gk-wc-karta"><?php the_content(); ?></div>
  <?php if ( is_cart() && function_exists( 'WC' ) && WC()->cart && WC()->cart->is_empty() ) : ?>
  <p class="gk-pusty"><a class="btn" href="<?php echo esc_url( home_url( '/materialy/' ) ); ?>">Zobacz materiały</a></p>
  <?php endif; ?>
</div></section>
		<?php
	endwhile;
endif;

echo $ch['foot']; // phpcs:ignore
?>
</div>
<?php wp_footer(); ?>
</body>
</html>
