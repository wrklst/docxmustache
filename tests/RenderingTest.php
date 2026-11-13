<?php

namespace WrkLst\DocxMustache\Tests;

use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\TestCase;
use WrkLst\DocxMustache\HtmlConversion;
use WrkLst\DocxMustache\MustacheRender;

/**
 * MustacheRender and HtmlConversion: plain string in, string out.
 */
class RenderingTest extends TestCase
{
    private const RUN = '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>';

    private const DATA = [
        'title'  => 'Tom & Jerry <"quoted">',
        'raw'    => 'raw &amp; ok',
        'html'   => '*[[DONOTESCAPE]]*<b>bold</b> & more',
        'number' => 42,
        'bool'   => true,
        'nul'    => null,
        'nested' => ['value' => 'deep'],
        'items'  => [['name' => 'A', 'pos' => 1], ['name' => 'Ümlaut', 'pos' => 2.5]],
        'many'   => [1, 2, 3, 4, 5, 6],
    ];

    #[DataProvider('templates')]
    public function test_render(string $template, string $expected): void
    {
        $this->assertSame($expected, MustacheRender::render(self::DATA, $template));
    }

    public static function templates(): array
    {
        return [
            'no tags'                => ['no tags at all', 'no tags at all'],
            'values are escaped'     => ['<w:t>{{title}}</w:t>', '<w:t>Tom &amp; Jerry &lt;&quot;quoted&quot;&gt;</w:t>'],
            'triple stays raw'       => ['<w:t>{{{raw}}} {{&raw}}</w:t>', '<w:t>raw &amp; ok raw &amp; ok</w:t>'],
            'DONOTESCAPE keeps html' => ['<w:t>{{html}}</w:t>', '<w:t><b>bold</b> &amp; more</w:t>'],
            'scalars and null'       => ['<w:t>{{number}}|{{bool}}|{{nul}}|{{missing}}|{{nested.value}}</w:t>', '<w:t>42|1|||deep</w:t>'],
            'sections'               => ['<w:t>{{#items}}[{{name}}:{{pos}}]{{/items}}{{^nothing}}empty{{/nothing}}</w:t>', '<w:t>[A:1][Ümlaut:2.5]empty</w:t>'],
            'array used as value'    => ['<w:t>{{nested}} {{many}}</w:t>', '<w:t>Array(value) Value used Array(0, 1, 2, 3, 4, ...) Value used</w:t>'],
            'tag split over runs'    => [
                '<w:r><w:t>{{ti</w:t></w:r><w:r><w:t>tle}}</w:t></w:r>',
                '<w:r><w:t xml:space="preserve">Tom &amp; Jerry &lt;&quot;quoted&quot;&gt;</w:t></w:r><w:r><w:t></w:t></w:r>',
            ],
        ];
    }

    public function test_split_tags_stay_split_without_cleaning(): void
    {
        $template = '<w:r><w:t>{{number</w:t></w:r><w:r><w:t>}}</w:t></w:r>';

        $this->assertStringContainsString('42', MustacheRender::render(self::DATA, $template));
        $this->assertStringNotContainsString('42', MustacheRender::render(self::DATA, $template, false));
    }

    #[DataProvider('splitTags')]
    public function test_tag_cleaner(string $template, string $expected): void
    {
        $this->assertSame($expected, MustacheRender::tagCleaner($template));
    }

    public static function splitTags(): array
    {
        return [
            'untouched without split tags' => ['<w:r><w:t>{{title}}</w:t></w:r><w:t>plain</w:t>', '<w:r><w:t>{{title}}</w:t></w:r><w:t>plain</w:t>'],
            'braces split from the name'   => [
                '<w:r><w:t>a {</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>{title}</w:t></w:r><w:r><w:t xml:space="preserve">} b</w:t></w:r>',
                '<w:r><w:t xml:space="preserve">a {{title}} b</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t></w:t></w:r><w:r><w:t xml:space="preserve"></w:t></w:r>',
            ],
            'xml between the runs is kept' => [
                '<w:r><w:t xml:space="preserve">x {{na</w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t>me</w:t></w:r><w:r><w:t>}}</w:t></w:r>',
                '<w:r><w:t xml:space="preserve">x {{name}}</w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t></w:t></w:r><w:r><w:t></w:t></w:r>',
            ],
            'single brace pulls in the next run' => ['<w:t>curly {</w:t><w:t>not a tag</w:t>', '<w:t xml:space="preserve">curly {not a tag</w:t><w:t></w:t>'],
        ];
    }

    #[DataProvider('html')]
    public function test_html_conversion(string $text, string $expected): void
    {
        $rPr = fn ($style = '') => '<w:r><w:rPr>'.$style.'<w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">';

        $this->assertSame(
            strtr($expected, ['[' => $rPr(), '[b:' => $rPr('<w:b />'), '[i:' => $rPr('<w:i />'), '[u:' => $rPr('<w:u w:val="single" />'), '[ib:' => $rPr('<w:i /><w:b />'), ']' => '</w:t></w:r>']),
            HtmlConversion::convert(self::RUN.$text.'</w:t></w:r>')
        );
    }

    /** Expected runs in shorthand: [text] is a plain run, [b:text] a bold one, and so on. */
    public static function html(): array
    {
        return [
            'plain text'   => ['nothing', '[nothing]'],
            'bold'         => ['x <b>bold</b> y', '[x ][b:bold][ y]'],
            'italic'       => ['x <i>it</i> y', '[x ][i:it][ y]'],
            'underline'    => ['x <u>un</u> y', '[x ][u:un][ y]'],
            'strong, em'   => ['<strong>s</strong> and <em>e</em>', '[][b:s][ and ][i:e][]'],
            'repeated'     => ['<b>1</b> mid <b>2</b>', '[][b:1][ mid ][b:2][]'],
            'nested'       => ['<b>bo <i>both</i> ld</b>', '[][b:bo ][ib:both][b: ld][]'],
            'line breaks'  => ['a<br />b<br/>c<br>d&lt;br /&gt;e&lt;br/&gt;f&lt;br&gt;g', '[a</w:t><w:br/><w:t xml:space="preserve">b</w:t><w:br/><w:t xml:space="preserve">c</w:t><w:br/><w:t xml:space="preserve">d</w:t><w:br/><w:t xml:space="preserve">e</w:t><w:br/><w:t xml:space="preserve">f</w:t><w:br/><w:t xml:space="preserve">g]'],
            'case matters' => ['<B>x</B>', '[<B>x</B>]'],
        ];
    }

    public function test_italic_overrides_an_explicit_non_italic_run(): void
    {
        $this->assertSame(
            '<w:r><w:rPr><w:i w:val="0"/></w:rPr><w:t xml:space="preserve"></w:t></w:r>'
            .'<w:r><w:rPr><w:i /></w:rPr><w:t xml:space="preserve">it</w:t></w:r>'
            .'<w:r><w:rPr><w:i w:val="0"/></w:rPr><w:t xml:space="preserve"></w:t></w:r>',
            HtmlConversion::convert('<w:r><w:rPr><w:i w:val="0"/></w:rPr><w:t><i>it</i></w:t></w:r>')
        );
    }

    public function test_unclosed_html_tag_does_not_raise(): void
    {
        $this->assertStringContainsString('never closed', HtmlConversion::convert(self::RUN.'x <b>never closed</w:t></w:r>'));
    }
}
