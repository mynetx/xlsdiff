<?php
/*
    Ross Scrivener http://scrivna.com
    PHP file diff implementation

    Much credit goes to...

    Paul's Simple Diff Algorithm v 0.1
    (C) Paul Butler 2007 <http://www.paulbutler.org/>
    May be used and distributed under the zlib/libpng license.

    ... for the actual diff code, i changed a few things and implemented a pretty interface to it.
*/
class diff {

    var $changes = array();
    var $diff = array();
    var $linepadding = null;

    function doDiff($old, $new){
        $maxlen = 0;

        if (!is_array($old)) $old = file($old);
        if (!is_array($new)) $new = file($new);

        foreach($old as $oindex => $ovalue){
            $nkeys = array_keys($new, $ovalue);
            foreach($nkeys as $nindex){
                $matrix[$oindex][$nindex] = isset($matrix[$oindex - 1][$nindex - 1]) ? $matrix[$oindex - 1][$nindex - 1] + 1 : 1;
                if($matrix[$oindex][$nindex] > $maxlen){
                    $maxlen = $matrix[$oindex][$nindex];
                    $omax = $oindex + 1 - $maxlen;
                    $nmax = $nindex + 1 - $maxlen;
                }
            }
        }
        if($maxlen == 0) return array(array('d'=>$old, 'i'=>$new));

        return array_merge(
                        $this->doDiff(array_slice($old, 0, $omax), array_slice($new, 0, $nmax)),
                        array_slice($new, $nmax, $maxlen),
                        $this->doDiff(array_slice($old, $omax + $maxlen), array_slice($new, $nmax + $maxlen)));

    }

    function diffWrap($old, $new){
        $this->diff = $this->doDiff($old, $new);
        $this->changes = array();
        $ndiff = array();
        foreach ($this->diff as $line => $k){
            if(is_array($k)){
                if (isset($k['d'][0]) || isset($k['i'][0])){
                    $this->changes[] = $line;
                    $ndiff[$line] = $k;
                }
            } else {
                $ndiff[$line] = $k;
            }
        }
        $this->diff = $ndiff;
        return $this->diff;
    }

    function formatcode($code){
        $code = trim($code);
        if (strlen($code) > 2) {
            $code = substr($code, 1, strlen($code) - 2);
        }
        $code = htmlentities($code);
        $code = str_replace(" ",'&nbsp;',$code);
        $code = str_replace("\t",'&nbsp;&nbsp;&nbsp;&nbsp;',$code);
        $code = str_replace('&quot;;&quot;','</td><td>',$code);
        return $code;
    }

    function showline($line){
        if ($this->linepadding === 0){
            if (in_array($line,$this->changes)) return true;
            return false;
        }
        if(is_null($this->linepadding)) return true;

        $start = (($line - $this->linepadding) > 0) ? ($line - $this->linepadding) : 0;
        $end = ($line + $this->linepadding);
        $search = range($start,$end);
        foreach($search as $k){
            if (in_array($k,$this->changes)) return true;
        }
        return false;

    }

    function inline($old, $new, $linepadding=null){
        $this->linepadding = $linepadding;

        $ret = '<pre><table width="100%" border="0" cellspacing="0" cellpadding="0" class="code">';
        $ret.= '<tr><td>' . STR_OLD . '</td><td>' . STR_NEW . '</td><td></td></tr>';
        $count_old = 1;
        $count_new = 1;

        $insert = false;
        $delete = false;
        $truncate = false;

        $diff = $this->diffWrap($old, $new);

        foreach($diff as $line => $k){
            if ($this->showline($line)){
                $truncate = false;
                if(is_array($k)){
                    foreach ($k['d'] as $val){
                        $class = '';
                        if (!$delete){
                            $delete = true;
                            $class = 'first';
                            if ($insert) $class = '';
                            $insert = false;
                        }
                        $ret.= '<tr class="del '.$class.'"><th>'.$count_old.'</th>';
                        $ret.= '<th>&nbsp;</th>';
                        $ret.= '<td>'.$this->formatcode($val).'</td>';
                        $ret.= '</tr>';
                        $count_old++;
                    }
                    foreach ($k['i'] as $val){
                        $class = '';
                        if (!$insert){
                            $insert = true;
                            $class = 'first';
                            if ($delete) $class = '';
                            $delete = false;
                        }
                        $ret.= '<tr class="ins '.$class.'"><th>&nbsp;</th>';
                        $ret.= '<th>'.$count_new.'</th>';
                        $ret.= '<td>'.$this->formatcode($val).'</td>';
                        $ret.= '</tr>';
                        $count_new++;
                    }
                } else {
                    $class = ($delete) ? 'del_end' : '';
                    $class = ($insert) ? 'ins_end' : $class;
                    $delete = false;
                    $insert = false;
                    $ret.= '<tr class="'.$class.'"><th>'.$count_old.'</th>';
                    $ret.= '<th>'.$count_new.'</th>';
                    $ret.= '<td>'.$this->formatcode($k).'</td>';
                    $ret.= '</tr>';
                    $count_old++;
                    $count_new++;
                }
            } else {
                $class = ($delete) ? 'del_end' : '';
                $class = ($insert) ? 'ins_end' : $class;
                $delete = false;
                $insert = false;

                if (!$truncate){
                    $truncate = true;
                    $ret.= '<tr class="truncated '.$class.'"><th>...</th>';
                    $ret.= '<th>...</th>';
                    $ret.= '<td>&nbsp;</td>';
                    $ret.= '</tr>';
                }
                $count_old++;
                $count_new++;

            }
        }
        $ret.= '</table></pre>';
        return $ret;
    }

    function style() {
        return <<< DIFF_STYLE
<style type="text/css">

table.code {
    border: 1px solid #ddd;
    border-spacing: 0;
    border-top: 0;
    empty-cells: show;
    font-size: 12px;
    line-height: 110%;
    padding: 0;
    margin: 0;
    width: 100%;
}

table.code th {
    background: #eed;
    color: #886;
    font-weight: normal;
    padding: 0 .5em;
    text-align: right;
    border-right: 1px solid #d7d7d7;
    border-top: 1px solid #998;
    font-size: 11px;
    width: 35px;
}

table.code td {
    background: #fff;
    font: normal 11px monospace;
    overflow: auto;
    padding: 1px 2px;
}

table.code tr.del td {
    background-color: #F99;
    color: #900;
}
table.code tr.del.first td {
    border-top: 1px solid #900;
}
table.code tr.ins td {
    background-color: #8f8;
}
table.code tr.ins.first td {
    border-top: 1px solid #090;
}
table.code tr.del_end td {
    border-top: 1px solid #900;
}
table.code tr.ins_end td {
    border-top: 1px solid #090;
}
table.code tr.truncated td {
    background-color: #f7f7f7;
}
</style>

DIFF_STYLE;
    }
}

$arguments = array();
if (isset($argv)) {
    $arguments = $argv;
}

if (count($arguments) < 6) {
    die();
}

define('STR_OLD', $arguments[4]);
define('STR_NEW', $arguments[5]);

$diff = new diff;
$text = $diff->style() . $diff->inline($arguments[1], $arguments[2], 2);
$fp = fopen($arguments[3], 'w');
fwrite($fp, $text);
fclose($fp);
