<?php


namespace Christophedlr\StMso\Excel;


class Exception extends \Exception
{
    public function __toString()
    {
        return sprintf("Excel Exception with code %d: %s", $this->code, $this->message);
    }
}
