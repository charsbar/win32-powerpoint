package Win32::PowerPoint::Constants;

use strict;
use Carp;

our $VERSION = '0.10';

our $AUTOLOAD;

sub new {
  my $class = shift;
  bless {

# ppSlideLayout
    ppLayoutBlank => 12,
    ppLayoutText  => 2,
    ppLayoutTitle => 1,

# ppAutoSize
    ppAutoSizeNone           => 0,
    ppAutoSizeShapeToFitText => 1,
    ppAutoSizeMixed          => -2,

# ppSaveAsFileType
    ppSaveAsPresentation => 1,
    ppSaveAsShow         => 7,

# ppParagraphAlignment
    ppAlignLeft       => 1,
    ppAlignCenter     => 2,
    ppAlignRight      => 3,
    ppAlignJustitfy   => 4,
    ppAlignDistribute => 5,
    ppAlignmentMixed  => -2,

# ppMouseActivation
    ppMouseClick => 1,
    ppMouseOver  => 2,

# ppDateTimeFormat
    ppDateTimeMdyy           => 1,
    ppDateTimeddddMMMMddyyyy => 2,
    ppDateTimedMMMMyyyy      => 3,
    ppDateTimeMMMMdyyyy      => 4,
    ppDateTimedMMMyy         => 5,
    ppDateTimeMMMMyy         => 6,
    ppDateTimeMMyy           => 7,
    ppDateTimeMMddyyHmm      => 8,
    ppDateTimeMMddyyhmmAMPM  => 9,
    ppDateTimeHmm            => 10,
    ppDateTimeHmmss          => 11,
    ppDateTimehmmAMPM        => 12,
    ppDateTimehmmssAMPM      => 13,
    ppDateTimeFormatMixed    => -2,
    
# msoAutoShapeType
    msoShape10pointStar                      => 149,
    msoShape12pointStar                      => 150,
    msoShape16pointStar                      => 94,
    msoShape24pointStar                      => 95,
    msoShape32pointStar                      => 96,
    msoShape4pointStar                       => 91,
    msoShape5pointStar                       => 92,
    msoShape6pointStar                       => 147,
    msoShape7pointStar                       => 148,
    msoShape8pointStar                       => 93,
    msoShapeActionButtonBackorPrevious       => 129,
    msoShapeActionButtonBeginning            => 131,
    msoShapeActionButtonCustom               => 125,
    msoShapeActionButtonDocument             => 134,
    msoShapeActionButtonEnd                  => 132,
    msoShapeActionButtonForwardorNext        => 130,
    msoShapeActionButtonHelp                 => 127,
    msoShapeActionButtonHome                 => 126,
    msoShapeActionButtonInformation          => 128,
    msoShapeActionButtonMovie                => 136,
    msoShapeActionButtonReturn               => 133,
    msoShapeActionButtonSound                => 135,
    msoShapeArc                              => 25,
    msoShapeBalloon                          => 137,
    msoShapeBentArrow                        => 41,
    msoShapeBentUpArrow                      => 44,
    msoShapeBevel                            => 15,
    msoShapeBlockArc                         => 20,
    msoShapeCan                              => 13,
    msoShapeChartPlus                        => 182,
    msoShapeChartStar                        => 181,
    msoShapeChartX                           => 180,
    msoShapeChevron                          => 52,
    msoShapeChord                            => 161,
    msoShapeCircularArrow                    => 60,
    msoShapeCloud                            => 179,
    msoShapeCloudCallout                     => 108,
    msoShapeCorner                           => 162,
    msoShapeCornerTabs                       => 169,
    msoShapeCross                            => 11,
    msoShapeCube                             => 14,
    msoShapeCurvedDownArrow                  => 48,
    msoShapeCurvedDownRibbon                 => 100,
    msoShapeCurvedLeftArrow                  => 46,
    msoShapeCurvedRightArrow                 => 45,
    msoShapeCurvedUpArrow                    => 47,
    msoShapeCurvedUpRibbon                   => 99,
    msoShapeDecagon                          => 144,
    msoShapeDiagonalStripe                   => 141,
    msoShapeDiamond                          => 4,
    msoShapeDodecagon                        => 146,
    msoShapeDonut                            => 18,
    msoShapeDoubleBrace                      => 27,
    msoShapeDoubleBracket                    => 26,
    msoShapeDoubleWave                       => 104,
    msoShapeDownArrow                        => 36,
    msoShapeDownArrowCallout                 => 56,
    msoShapeDownRibbon                       => 98,
    msoShapeExplosion1                       => 89,
    msoShapeExplosion2                       => 90,
    msoShapeFlowchartAlternateProcess        => 62,
    msoShapeFlowchartCard                    => 75,
    msoShapeFlowchartCollate                 => 79,
    msoShapeFlowchartConnector               => 73,
    msoShapeFlowchartData                    => 64,
    msoShapeFlowchartDecision                => 63,
    msoShapeFlowchartDelay                   => 84,
    msoShapeFlowchartDirectAccessStorage     => 87,
    msoShapeFlowchartDisplay                 => 88,
    msoShapeFlowchartDocument                => 67,
    msoShapeFlowchartExtract                 => 81,
    msoShapeFlowchartInternalStorage         => 66,
    msoShapeFlowchartMagneticDisk            => 86,
    msoShapeFlowchartManualInput             => 71,
    msoShapeFlowchartManualOperation         => 72,
    msoShapeFlowchartMerge                   => 82,
    msoShapeFlowchartMultidocument           => 68,
    msoShapeFlowchartOfflineStorage          => 139,
    msoShapeFlowchartOffpageConnector        => 74,
    msoShapeFlowchartOr                      => 78,
    msoShapeFlowchartPredefinedProcess       => 65,
    msoShapeFlowchartPreparation             => 70,
    msoShapeFlowchartProcess                 => 61,
    msoShapeFlowchartPunchedTape             => 76,
    msoShapeFlowchartSequentialAccessStorage => 85,
    msoShapeFlowchartSort                    => 80,
    msoShapeFlowchartStoredData              => 83,
    msoShapeFlowchartSummingJunction         => 77,
    msoShapeFlowchartTerminator              => 69,
    msoShapeFoldedCorner                     => 16,
    msoShapeFrame                            => 158,
    msoShapeFunnel                           => 174,
    msoShapeGear6                            => 172,
    msoShapeGear9                            => 173,
    msoShapeHalfFrame                        => 159,
    msoShapeHeart                            => 21,
    msoShapeHeptagon                         => 145,
    msoShapeHexagon                          => 10,
    msoShapeHorizontalScroll                 => 102,
    msoShapeIsoscelesTriangle                => 7,
    msoShapeLeftArrow                        => 34,
    msoShapeLeftArrowCallout                 => 54,
    msoShapeLeftBrace                        => 31,
    msoShapeLeftBracket                      => 29,
    msoShapeLeftCircularArrow                => 176,
    msoShapeLeftRightArrow                   => 37,
    msoShapeLeftRightArrowCallout            => 57,
    msoShapeLeftRightCircularArrow           => 177,
    msoShapeLeftRightRibbon                  => 140,
    msoShapeLeftRightUpArrow                 => 40,
    msoShapeLeftUpArrow                      => 43,
    msoShapeLightningBolt                    => 22,
    msoShapeLineCallout1                     => 109,
    msoShapeLineCallout1AccentBar            => 113,
    msoShapeLineCallout1BorderandAccentBar   => 121,
    msoShapeLineCallout1NoBorder             => 117,
    msoShapeLineCallout2                     => 110,
    msoShapeLineCallout2AccentBar            => 114,
    msoShapeLineCallout2BorderandAccentBar   => 122,
    msoShapeLineCallout2NoBorder             => 118,
    msoShapeLineCallout3                     => 111,
    msoShapeLineCallout3AccentBar            => 115,
    msoShapeLineCallout3BorderandAccentBar   => 123,
    msoShapeLineCallout3NoBorder             => 119,
    msoShapeLineCallout4                     => 112,
    msoShapeLineCallout4AccentBar            => 116,
    msoShapeLineCallout4BorderandAccentBar   => 124,
    msoShapeLineCallout4NoBorder             => 120,
    msoShapeLineInverse                      => 183,
    msoShapeMathDivide                       => 166,
    msoShapeMathEqual                        => 167,
    msoShapeMathMinus                        => 164,
    msoShapeMathMultiply                     => 165,
    msoShapeMathNotEqual                     => 168,
    msoShapeMathPlus                         => 163,
    msoShapeMixed                            => -2,
    msoShapeMoon                             => 24,
    msoShapeNoSymbol                         => 19,
    msoShapeNonIsoscelesTrapezoid            => 143,
    msoShapeNotPrimitive                     => 138,
    msoShapeNotchedRightArrow                => 50,
    msoShapeOctagon                          => 6,
    msoShapeOval                             => 9,
    msoShapeOvalCallout                      => 107,
    msoShapeParallelogram                    => 2,
    msoShapePentagon                         => 51,
    msoShapePie                              => 142,
    msoShapePieWedge                         => 175,
    msoShapePlaque                           => 28,
    msoShapePlaqueTabs                       => 171,
    msoShapeQuadArrow                        => 39,
    msoShapeQuadArrowCallout                 => 59,
    msoShapeRectangle                        => 1,
    msoShapeRectangularCallout               => 105,
    msoShapeRegularPentagon                  => 12,
    msoShapeRightArrow                       => 33,
    msoShapeRightArrowCallout                => 53,
    msoShapeRightBrace                       => 32,
    msoShapeRightBracket                     => 30,
    msoShapeRightTriangle                    => 8,
    msoShapeRound1Rectangle                  => 151,
    msoShapeRound2DiagRectangle              => 153,
    msoShapeRound2SameRectangle              => 152,
    msoShapeRoundedRectangle                 => 5,
    msoShapeRoundedRectangularCallout        => 106,
    msoShapeSmileyFace                       => 17,
    msoShapeSnip1Rectangle                   => 155,
    msoShapeSnip2DiagRectangle               => 157,
    msoShapeSnip2SameRectangle               => 156,
    msoShapeSnipRoundRectangle               => 154,
    msoShapeSquareTabs                       => 170,
    msoShapeStripedRightArrow                => 49,
    msoShapeSun                              => 23,
    msoShapeSwooshArrow                      => 178,
    msoShapeTear                             => 160,
    msoShapeTrapezoid                        => 3,
    msoShapeUTurnArrow                       => 42,
    msoShapeUpArrow                          => 35,
    msoShapeUpArrowCallout                   => 55,
    msoShapeUpDownArrow                      => 38,
    msoShapeUpDownArrowCallout               => 58,
    msoShapeUpRibbon                         => 97,
    msoShapeVerticalScroll                   => 101,
    msoShapeWave                             => 103,

# msoPatternType
    msoPattern10Percent              => 2,
    msoPattern20Percent              => 3,
    msoPattern25Percent              => 4,
    msoPattern30Percent              => 5,
    msoPattern40Percent              => 6,
    msoPattern50Percent              => 7,
    msoPattern5Percent               => 1,
    msoPattern60Percent              => 8,
    msoPattern70Percent              => 9,
    msoPattern75Percent              => 10,
    msoPattern80Percent              => 11,
    msoPattern90Percent              => 12,
    msoPatternDarkDownwardDiagonal   => 15,
    msoPatternDarkHorizontal         => 13,
    msoPatternDarkUpwardDiagonal     => 16,
    msoPatternDarkVertical           => 14,
    msoPatternDashedDownwardDiagonal => 28,
    msoPatternDashedHorizontal       => 32,
    msoPatternDashedUpwardDiagonal   => 27,
    msoPatternDashedVertical         => 31,
    msoPatternDiagonalBrick          => 40,
    msoPatternDivot                  => 46,
    msoPatternDottedDiamond          => 24,
    msoPatternDottedGrid             => 45,
    msoPatternHorizontalBrick        => 35,
    msoPatternLargeCheckerBoard      => 36,
    msoPatternLargeConfetti          => 33,
    msoPatternLargeGrid              => 34,
    msoPatternLightDownwardDiagonal  => 21,
    msoPatternLightHorizontal        => 19,
    msoPatternLightUpwardDiagonal    => 22,
    msoPatternLightVertical          => 20,
    msoPatternMixed                  => -2,
    msoPatternNarrowHorizontal       => 30,
    msoPatternNarrowVertical         => 29,
    msoPatternOutlinedDiamond        => 41,
    msoPatternPlaid                  => 42,
    msoPatternShingle                => 47,
    msoPatternSmallCheckerBoard      => 17,
    msoPatternSmallConfetti          => 37,
    msoPatternSmallGrid              => 23,
    msoPatternSolidDiamond           => 39,
    msoPatternSphere                 => 43,
    msoPatternTrellis                => 18,
    msoPatternWave                   => 48,
    msoPatternWeave                  => 44,
    msoPatternWideDownwardDiagonal   => 25,
    msoPatternWideUpwardDiagonal     => 26,
    msoPatternZigZag                 => 38,

# msoTextOrientation
    msoTextOrientationHorizontal => 1,

# msoTriState
    msoTrue  => -1,
    msoFalse => 0,

  }, $class;
}

sub AUTOLOAD {
  my $self = shift;
  my $name = $AUTOLOAD;
  $name =~ s/.*://;
  if (exists $self->{$name})      { return $self->{$name}; }
  if (exists $self->{"pp$name"})  { return $self->{"pp$name"}; }
  if (exists $self->{"mso$name"}) { return $self->{"mso$name"}; }
  croak "constant $name does not exist";
}

sub DESTROY {}

1;
__END__

=head1 NAME

Win32::PowerPoint::Constants - Constants holder

=head1 DESCRIPTION

This is used internally in L<Win32::PowerPoint>.

=head1 METHOD

=head2 new

Creates an object.

=head1 SEE ALSO

PowerPoint's object browser and MSDN documentation.

=head1 AUTHOR

Kenichi Ishigaki, E<lt>ishigaki@cpan.orgE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2006 by Kenichi Ishigaki

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut
