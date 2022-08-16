package com.example.exceldemo

import android.app.Activity
import android.content.Intent
import android.os.Bundle
import android.os.Environment
import android.view.View
import androidx.activity.result.contract.ActivityResultContracts.StartActivityForResult
import androidx.appcompat.app.AppCompatActivity
import androidx.recyclerview.widget.LinearLayoutManager
import com.example.exceldemo.databinding.ActivityListBinding
import com.google.android.material.snackbar.Snackbar
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.usermodel.HSSFWorkbookFactory
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream


class ListActivity : AppCompatActivity() {
    private lateinit var mAdapter: MyAdapter
    private var list = mutableListOf<UserModel>()
    private var dir = File(
        Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DCIM).absolutePath.plus(
            "/Demo5.xls"
        )
    )
    private lateinit var editer: HSSFWorkbook

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        val binding = ActivityListBinding.inflate(layoutInflater)
        val view: View = binding.root
        setContentView(view)
        try {
            if (!dir.exists()) {
//                dir.createNewFile()
                editer = HSSFWorkbook()
                val fileOutputStream = FileOutputStream(dir)
                editer.write(fileOutputStream)
                fileOutputStream.close()
            } else {
                val inputStream = FileInputStream(dir)
                editer = HSSFWorkbookFactory.create(inputStream) as HSSFWorkbook
                inputStream.close()
                list = editer.getSheet("Sheet0").toMutableList().map {
                    UserModel(
                        it.getCell(0).stringCellValue,
                        it.getCell(1).stringCellValue,
                        it.getCell(2).stringCellValue,
                        it.getCell(3).stringCellValue,
                        it.getCell(4).stringCellValue,
                        it.getCell(5).stringCellValue,
                        it.getCell(6).stringCellValue,
                        it.getCell(7).stringCellValue,
                        it.getCell(8).stringCellValue,
                    )
                }.toMutableList()
            }

        } catch (_: java.lang.Exception) {

        }

        val resultEditLauncher = registerForActivityResult(StartActivityForResult()) { result ->
            if (result.resultCode == Activity.RESULT_OK) {
                // There are no request codes
                val data: Intent? = result.data
                val model = data?.extras?.getSerializable("MODEL") as UserModel
                list[model.position ?: 0] = model
                mAdapter.notifyItemChanged(model.position ?: 0)
                Snackbar.make(
                    this.window.decorView.findViewById(android.R.id.content),
                    "Đã sửa",
                    Snackbar.LENGTH_SHORT
                ).show()
            }
        }

        mAdapter = MyAdapter(list)

        val layoutManager = LinearLayoutManager(this, LinearLayoutManager.VERTICAL, false)
        binding.rvList.layoutManager = layoutManager
        binding.rvList.adapter = mAdapter
        mAdapter.onClickItem = { position ->
            val obj = list[position].apply {
                this.position = position
            }
            val intent = Intent(this, DetailActivity::class.java).apply {
                putExtra("USER_DATA", obj)
            }
            resultEditLauncher.launch(intent)
        }


        val resultAddLauncher = registerForActivityResult(StartActivityForResult()) { result ->
            if (result.resultCode == Activity.RESULT_OK) {
                // There are no request codes
                val data: Intent? = result.data
                list.add(data?.extras?.getSerializable("MODEL") as UserModel)
                mAdapter.notifyDataSetChanged()
                Snackbar.make(
                    this.window.decorView.findViewById(android.R.id.content),
                    "Đã lưu",
                    Snackbar.LENGTH_SHORT
                ).show()
            }
        }

        binding.btnAdd.setOnClickListener {
            val intent = Intent(this, DetailActivity::class.java).apply {
                putExtra("SIZE", list.size)
            }
            resultAddLauncher.launch(intent)
        }


//        val sheet = editer.getSheet("list")
//        Log.e("Phucdz", sheet.getRow(0).getCell(0).stringCellValue)
//        Log.e("size", editer.getSheet("list").toList().size.toString())
    }
}