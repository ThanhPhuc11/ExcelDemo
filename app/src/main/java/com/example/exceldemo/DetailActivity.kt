package com.example.exceldemo

import android.Manifest
import android.app.Activity
import android.content.Intent
import android.content.pm.PackageManager
import android.os.Bundle
import android.os.Environment
import android.text.TextUtils
import android.view.View
import android.widget.ArrayAdapter
import android.widget.AutoCompleteTextView
import android.widget.EditText
import android.widget.TextView
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import com.example.exceldemo.databinding.ActivityMainBinding
import org.apache.poi.hssf.usermodel.HSSFWorkbookFactory
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream


class DetailActivity : AppCompatActivity() {
    private var dir = File(
        Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DCIM).absolutePath.plus(
            "/Demo5.xls"
        )
    )

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        val binding = ActivityMainBinding.inflate(layoutInflater)
        val view: View = binding.root
        setContentView(view)
        if (!hasPermission()) {
            startLocationPermissionRequest()
        }
        val listSize = intent.extras?.getInt("SIZE", 0) ?: 0
        val objEdit = try {
            intent?.extras?.getSerializable("USER_DATA") as UserModel
        } catch (_: Exception) {
            null
        }
        objEdit?.let { fillData(it, binding) }

        initViewSpinner(mutableListOf("V1", "NH", "LH", "V2"), binding.edtLoaiDichVu)
        initViewSpinner(mutableListOf("Thu ngay", "Thu sau"), binding.edtHinhThucThu)
        initViewSpinner(mutableListOf("C1", "C2"), binding.edtNoiSatHach)

        findViewById<TextView>(R.id.tvCreate).setOnClickListener {
            if (!isValidateName(binding.edtTenKH)) {
                Toast.makeText(this, "Chưa nhập tên khách hàng", Toast.LENGTH_SHORT).show()
                binding.edtTenKH.requestFocus()
                return@setOnClickListener
            }
            val inputStream = FileInputStream(dir)
            val editer = HSSFWorkbookFactory.create(inputStream)
            val sheet = try {
                editer.getSheetAt(0)
            } catch (_: Exception) {
                editer.createSheet("Sheet0")
            }
            val row = if (objEdit == null) sheet.createRow(listSize) else sheet.getRow(
                objEdit.position ?: 0
            )

            row.createCell(0).setCellValue(binding.edtTenKH.text.toString())
            row.createCell(1).setCellValue(binding.edtSdt.text.toString())
            row.createCell(2).setCellValue(binding.edtNamSinh.text.toString())
            row.createCell(3).setCellValue(binding.edtDiaChi.text.toString())
            row.createCell(4).setCellValue(binding.edtLoaiDichVu.text.toString())
            row.createCell(5).setCellValue(binding.edtHinhThucThu.text.toString())
            row.createCell(6).setCellValue(binding.edtNgaySatHach.text.toString())
            row.createCell(7).setCellValue(binding.edtNoiSatHach.text.toString())
            row.createCell(8).setCellValue(binding.edtGhiChu.text.toString())

//            if (!dir.exists()) {
//                dir.createNewFile()
//            }
            val fileOutputStream = FileOutputStream(dir)
            editer.write(fileOutputStream)
            val userModel = UserModel(
                row.getCell(0).stringCellValue,
                row.getCell(1).stringCellValue,
                row.getCell(2).stringCellValue,
                row.getCell(3).stringCellValue,
                row.getCell(4).stringCellValue,
                row.getCell(5).stringCellValue,
                row.getCell(6).stringCellValue,
                row.getCell(7).stringCellValue,
                row.getCell(8).stringCellValue,
                objEdit?.position
            )
            setResult(
                Activity.RESULT_OK,
                Intent().apply {
                    putExtra("MODEL", userModel)
                })


//            Toast.makeText(this, "Đã lưu", Toast.LENGTH_SHORT).show()
            finish()
        }

    }

    private val requestPermission =
        registerForActivityResult(ActivityResultContracts.RequestPermission()) {
            if (it) {
                Toast.makeText(this, "OK", Toast.LENGTH_SHORT).show()
            } else {
                Toast.makeText(this, "Ko cấp quyền tí nó crash chết mịa", Toast.LENGTH_SHORT).show()
            }
        }

    private fun startLocationPermissionRequest() {
        requestPermission.launch(Manifest.permission.WRITE_EXTERNAL_STORAGE)
    }

    private fun hasPermission(): Boolean =
        ActivityCompat.checkSelfPermission(
            this,
            Manifest.permission.WRITE_EXTERNAL_STORAGE
        ) == PackageManager.PERMISSION_GRANTED

    private fun initViewSpinner(list: List<String>, view: AutoCompleteTextView) {

        val adapter = ArrayAdapter(
            this,
            android.R.layout.simple_spinner_dropdown_item,
            list
        )

        view.setAdapter(adapter)
    }

    private fun isValidateName(view: EditText): Boolean {
        return view.text.isNotBlank()
    }

    private fun fillData(obj: UserModel, binding: ActivityMainBinding) {
        binding.edtTenKH.setText(obj.name)
        binding.edtSdt.setText(obj.phone)
        binding.edtNamSinh.setText(obj.age)
        binding.edtDiaChi.setText(obj.address)
        binding.edtLoaiDichVu.setText(obj.loaiDichVu)
        binding.edtHinhThucThu.setText(obj.hinhThucThu)
        binding.edtNgaySatHach.setText(obj.ngaySatHach)
        binding.edtNoiSatHach.setText(obj.noiSatHach)
        binding.edtGhiChu.setText(obj.ghiChu)
    }
}