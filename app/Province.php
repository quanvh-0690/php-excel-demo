<?php

namespace App;

use Illuminate\Database\Eloquent\Model;

class Province extends Model
{
    protected $fillable = [
        'id',
        'name',
        'type',
    ];
    
    public function districts()
    {
        return $this->hasMany(District::class);
    }
}
