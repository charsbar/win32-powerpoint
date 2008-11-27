#!perl
use strict;
use warnings;
use Win32::PowerPoint;

my $pp = Win32::PowerPoint->new;
$pp->new_presentation;

open my $fh, '<', shift or die $!;
my $new_slide = 1;
my $text     = '';
my $subtitle = '';
while( <$fh> ) {
  chomp;
  if ( $new_slide ) {
    if ( $text or $subtitle ) {
      $text     =~ s/\n+$//;
      $subtitle =~ s/\n+$//;
      $pp->add_text( $text, {
        top    => 50,
        left   => 50,
        width  => 620,
        height => 200,
        align  => 'center',
        size   => 72,
        font   => 'Comic Sans MS',
      }) if $text;
      $pp->add_text( $subtitle, {
        top    => 450,
        left   => 50,
        width  => 620,
        height => 50,
        align  => 'center',
        size   => 32,
        bold   => 1,
      }) if $subtitle;
    }
    $pp->new_slide;
    $new_slide = 0;
    $text = $subtitle = '';
  }
  if ( /^### (.+)/ ) {
    $subtitle .= "$1\n";
    next;
  }
  elsif ( /^\-\-\-\-/ ) {
    $new_slide = 1;
    next;
  }
  $text .= "$_\n";
}
close $fh;

$pp->save_presentation('slides.ppt');
