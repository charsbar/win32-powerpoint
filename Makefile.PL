use strict;
use warnings;
use ExtUtils::MakeMaker;

WriteMakefile(
    NAME                => 'Win32::PowerPoint',
    AUTHOR              => 'Kenichi Ishigaki <ishigaki@cpan.org>',
    VERSION_FROM        => 'lib/Win32/PowerPoint.pm',
    ABSTRACT_FROM       => 'lib/Win32/PowerPoint.pm',
    PREREQ_PM => {
        'Test::More'      => 0,
        'Win32::OLE'      => 0,
        'File::Spec'      => 0,
        'Exporter::Lite'  => 0,
        'Class::Accessor' => 0,
    },
);
