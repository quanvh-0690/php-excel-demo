<?php

namespace App;

use Illuminate\Database\Eloquent\Model;

class District extends Model
{
    protected $fillable = [
        'id',
        'name',
        'type',
        'location',
        'province_id',
    ];
    
    public function wards()
    {
        return $this->hasMany(Ward::class);
    }
}
