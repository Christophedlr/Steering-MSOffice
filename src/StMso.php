<?php


namespace Christophedlr\StMso;


use Christophedlr\StMso\Excel\Excel;

class StMso
{
    private $com;

    public function __construct()
    {
    }

    public function excel()
    {
        return new Excel();
    }
}
