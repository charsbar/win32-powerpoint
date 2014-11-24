use strict;
use warnings;
use ExtUtils::MakeMaker;

WriteMakefile(
  NAME                => 'Win32::PowerPoint',
  AUTHOR              => 'Kenichi Ishigaki <ishigaki@cpan.org>',
  VERSION_FROM        => 'lib/Win32/PowerPoint.pm',
  ABSTRACT_FROM       => 'lib/Win32/PowerPoint.pm',
  PREREQ_PM => {
    'Exporter::Lite'      => 0,
    'File::Spec'          => 0,
    'Test::More'          => '0.47',
    'Test::UseAllModules' => '0.09',
    'Win32::OLE'          => 0,

    (( $^O =~ /cygwin/ && $] < 5.010 )
      ? ( 'Filesys::CygwinPaths' => 0 )
      : ()
    ),

  },
  ($ExtUtils::MakeMaker::VERSION >= 6.31 ?
    (LICENSE => 'perl') : ()
  ),
);
