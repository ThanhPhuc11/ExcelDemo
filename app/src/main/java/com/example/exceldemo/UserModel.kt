package com.example.exceldemo

data class UserModel(
    var name: String? = null,
    var phone: String? = null,
    var age: String? = null,
    var address: String? = null,
    var loaiDichVu: String? = null,
    var hinhThucThu: String? = null,
    var ngaySatHach: String? = null,
    var noiSatHach: String? = null,
    var ghiChu: String? = null,
    var position: Int? = null
): java.io.Serializable
