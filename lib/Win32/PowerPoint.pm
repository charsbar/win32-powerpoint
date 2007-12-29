package Win32::PowerPoint;

use strict;
use warnings;
use Carp;

our $VERSION = '0.06';

use File::Spec;
use Win32::OLE;
use Win32::PowerPoint::Constants;
use Win32::PowerPoint::Utils qw( RGB canonical_alignment canonical_pattern );

use base qw( Class::Accessor::Fast );

__PACKAGE__->mk_ro_accessors( qw( c application presentation slide ) );

sub new {
  my $class = shift;
  my $self  = bless {
    c            => Win32::PowerPoint::Constants->new,
    was_invoked  => 0,
    application  => undef,
    presentation => undef,
    slide        => undef,
  }, $class;

  $self->connect_or_invoke;

  return $self;
}

##### application #####

sub connect_or_invoke {
  my $self = shift;

  $self->{application} = Win32::OLE->GetActiveObject('PowerPoint.Application');

  unless (defined $self->{application}) {
    $self->{application} = Win32::OLE->new('PowerPoint.Application')
      or die Win32::OLE->LastError;
    $self->{was_invoked} = 1;
  }
}

sub quit {
  my $self = shift;

  return unless $self->application;

  $self->application->Quit;
  $self->{application} = undef;
}

##### presentation #####

sub new_presentation {
  my $self = shift;

  return unless $self->{application};

  my %options = ( @_ == 1 and ref $_[0] eq 'HASH' ) ? %{ $_[0] } : @_;

  $self->{slide} = undef;

  $self->{presentation} = $self->application->Presentations->Add
    or die Win32::OLE->LastError;

  if ( my $color = $options{background_forecolor} || $options{masterbkgforecolor} ) {
    $self->presentation->SlideMaster->Background->Fill->ForeColor->{RGB} = RGB($color);
  }

  if ( my $color = $options{background_backcolor} || $options{masterbkgbackcolor} ) {
    $self->presentation->SlideMaster->Background->Fill->BackColor->{RGB} = RGB($color);
  }

  if ( $options{pattern} ) {
    my $method = canonical_pattern($options{pattern});
    $self->presentation->SlideMaster->Background->Fill->Patterned( $self->c->$method );
  }
}

sub save_presentation {
  my ($self, $file) = @_;

  return unless $self->presentation;

  $self->presentation->SaveAs( File::Spec->rel2abs($file) );
}

sub close_presentation {
  my $self = shift;

  $self->presentation->Close;
  $self->{presentation} = undef;
}

##### slide #####

sub new_slide {
  my $self = shift;

  my %options = ( @_ == 1 and ref $_[0] eq 'HASH' ) ? %{ $_[0] } : @_;

  $self->{slide} = $self->presentation->Slides->Add(
    $self->presentation->Slides->Count + 1,
    $self->c->LayoutBlank
  ) or die Win32::OLE->LastError;

  if ( my $color = $options{background_forecolor} || $options{bkgforecolor} ) {
    $self->slide->{FollowMasterBackground} = $self->c->msoFalse;
    $self->slide->Background->Fill->ForeColor->{RGB} = RGB($color);
  }

  if ( my $color = $options{background_backcolor} || $options{bkgbackcolor} ) {
    $self->slide->{FollowMasterBackground} = $self->c->msoFalse;
    $self->slide->Background->Fill->BackColor->{RGB} = RGB($color);
  }

  if ( $options{pattern} ) {
    my $method = canonical_pattern($options{pattern});
    $self->slide->Background->Fill->Patterned($self->c->$method);
  }
}

sub add_text {
  my ($self, $text, $options) = @_;

  return unless $self->slide;

  $options = {} unless ref $options eq 'HASH';

  $text =~ s/\n/\r/gs;

  my $num_of_boxes = $self->slide->Shapes->Count;
  my $last  = $num_of_boxes ? $self->slide->Shapes($num_of_boxes) : undef;
  my ($left, $top, $width, $height);
  if ($last) {
    $left   = $options->{left}   || $last->Left;
    $top    = $options->{top}    || $last->Top + $last->Height + 20;
    $width  = $options->{width}  || $last->Width;
    $height = $options->{height} || $last->Height;
  }
  else {
    $left   = $options->{left}   || 30;
    $top    = $options->{top}    || 30;
    $width  = $options->{width}  || 600;
    $height = $options->{height} || 200;
  }

  my $new_textbox = $self->slide->Shapes->AddTextbox(
    $self->c->TextOrientationHorizontal,
    $left, $top, $width, $height
  );

  my $frame = $new_textbox->TextFrame;
  my $range = $frame->TextRange;

  $frame->{WordWrap} = $self->c->True;
  $range->ParagraphFormat->{FarEastLineBreakControl} = $self->c->True;
  $range->{Text} = $text;

  $self->decorate_range( $range, $options );

  $frame->{AutoSize} = $self->c->AutoSizeNone;
  $frame->{AutoSize} = $self->c->AutoSizeShapeToFitText;

  return $new_textbox;
}

sub insert_before {
  my ($self, $text, $options) = @_;

  return unless $self->slide;

  $options = {} unless ref $options eq 'HASH';

  $text =~ s/\n/\r/gs;

  my $num_of_boxes = $self->slide->Shapes->Count;
  my $last  = $num_of_boxes ? $self->slide->Shapes($num_of_boxes) : undef;
  my $range = $self->slide->Shapes($num_of_boxes)->TextFrame->TextRange;

  my $selection = $range->InsertBefore($text);

  $self->decorate_range( $selection, $options );

  return $selection;
}

sub insert_after {
  my ($self, $text, $options) = @_;

  return unless $self->slide;

  $options = {} unless ref $options eq 'HASH';

  $text =~ s/\n/\r/gs;

  my $num_of_boxes = $self->slide->Shapes->Count;
  my $last  = $num_of_boxes ? $self->slide->Shapes($num_of_boxes) : undef;
  my $range = $self->{slide}->Shapes($num_of_boxes)->TextFrame->TextRange;

  my $selection = $range->InsertAfter($text);

  $self->decorate_range( $selection, $options );

  return $selection;
}

sub decorate_range {
  my ($self, $range, $options) = @_;

  my ($true, $false) = ($self->c->True, $self->c->False);

  $range->Font->{Bold}        = $options->{bold}        ? $true : $false;
  $range->Font->{Italic}      = $options->{italic}      ? $true : $false;
  $range->Font->{Underline}   = $options->{underline}   ? $true : $false;
  $range->Font->{Shadow}      = $options->{shadow}      ? $true : $false;
  $range->Font->{Subscript}   = $options->{subscript}   ? $true : $false;
  $range->Font->{Superscript} = $options->{superscript} ? $true : $false;
  $range->Font->{Size}        = $options->{size}       if $options->{size};
  $range->Font->{Name}        = $options->{name}       if $options->{name};
  $range->Font->{Name}        = $options->{font}       if $options->{font};
  $range->Font->Color->{RGB}  = RGB($options->{color}) if $options->{color};

  my $method = canonical_alignment( $options->{alignment} || $options->{align} || 'left' );
  $range->ParagraphFormat->{Alignment} = $self->c->$method;

  $range->ActionSettings(
    $self->c->MouseClick
  )->Hyperlink->{Address} = $options->{link} if $options->{link};
}

sub DESTROY {
  my $self = shift;

  $self->quit if $self->{was_invoked};
}

1;
__END__

=head1 NAME

Win32::PowerPoint - helps to convert texts to PP slides

=head1 SYNOPSIS

    use Win32::PowerPoint;

    # invoke (or connect to) PowerPoint
    my $pp = Win32::PowerPoint->new;

    $pp->new_presentation(
      background_forecolor => [255,255,255],
      background_backcolor => 'RGB(0, 0, 0)',
      pattern => 'Shingle',
    );

    ... (load and parse your slide text)

    foreach my $slide (@slides) {
      $pp->new_slide;

      $pp->add_text($slide->title, { size => 40, bold => 1 });
      $pp->add_text($slide->body);
      $pp->add_text($slide->link,  { link => $slide->link });
    }

    $pp->save_presentation('slide.ppt');

    $pp->close_presentation;

    # PowerPoint closes automatically

=head1 DESCRIPTION

Win32::PowerPoint mainly aims to help to convert L<Spork> (or Sporx)
texts to PowerPoint slides. Though there's no converter at the moment,
you can add texts to your new slides/presentation and save it. 

=head1 METHODS

=head2 new

Invokes (or connects to) PowerPoint.

=head2 connect_or_invoke

Explicitly connects to (or invoke) PowerPoint.

=head2 quit

Explicitly disconnects and close PowerPoint this module (or you) invoked.

=head2 new_presentation (options)

Creates a new (probably blank) presentation. Options are:

=over 4

=item background_forecolor, background_backcolor

You can specify background colors of the slides with an array ref of RGB
components ([255, 255, 255] for white) or formatted string ('255, 0, 0'
for red). You can use '(0, 255, 255)' or 'RGB(0, 255, 255)' format for
clarity. These colors are applied to all the slides you'll add, unless
you specify other colors for the slides explicitly.

You can use 'masterbkgforecolor' and 'masterbkgbackcolor' as aliases.

=item pattern

You also can specify default background pattern for the slides.
See L<Win32::PowerPoint::Constants> (or MSDN or PowerPoint's help) for
supported pattern names. You can omit 'msoPattern' part and the names
are case-sensitive.

=back

=head2 save_presentation (path)

Saves the presentation to where you specified. Accepts relative path.
You might want to save it as .pps (slideshow) file to make it easy to
show slides (it just starts full screen slideshow with doubleclick).

=head2 close_presentation

Explicitly closes the presentation.

=head2 new_slide (options)

Adds a new (blank) slide to the presentation. Options are:

=over 4

=item background_forecolor, background_backcolor

You can set colors just for the slide with these options.
You can use 'bkgforecolor' and 'bkgbackcolor' as aliases.

=item pattern

You also can set background pattern just for the slide.

=back

=head2 add_text (text, options)

Adds (formatted) text to the slide. Options are:

=over 4

=item left, top, width, height

of the Textbox.

=back

See 'decorate_range' for other options.

=head2 insert_before (text, options)

=head2 insert_after (text, options)

Prepends/Appends text to the current Textbox. See 'decorate_range' for options.

=head2 decorate_range (range, options)

Decorates text of the range. Options are:

=over 4

=item bold, italic, underline, shadow, subscript, superscript

Boolean.

=item size

Integer.

=item color

See above for the convention.

=item font

Font name of the text. You can use 'name' as an alias.

=item alignment

One of the 'left' (default), 'center', 'right', 'justify', 'distribute'.

You can use 'align' as an alias.

=item link

hyperlink address of the Text.

=back

(This method is mainly for the internal use).

=head1 IF YOU WANT TO GO INTO DETAIL

This module uses Win32::OLE internally. You can fully control PowerPoint
through these accessors.

=head2 application

returns Application object.

    print $pp->application->Name;

=head2 presentation

returns current Presentation object (maybe ActivePresentation but that's
not assured).

    $pp->save_presentation('sample.ppt') unless $pp->presentation->Saved;

    while (my $last = $pp->presentation->Slides->Count) {
      $pp->presentation->Slides($last)->Delete;
    }

=head2 slide

returns current Slide object.

    $pp->slide->Export(".\\slide_01.jpg",'jpg');

    $pp->slide->Shapes(1)->TextFrame->TextRange
       ->Characters(1, 5)->Font->{Bold} = $pp->c->True;

=head2 c

returns Win32::PowerPoint::Constants object.

=head1 AUTHOR

Kenichi Ishigaki, E<lt>ishigaki@cpan.orgE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2006- by Kenichi Ishigaki

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut

