package com.example.nexttry

import android.content.Intent
import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import com.example.nexttry.databinding.ActivityMainBinding
import java.io.File
import android.util.Log
import androidx.core.content.FileProvider

class MainActivity : AppCompatActivity() {
    private lateinit var binding: ActivityMainBinding

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        binding.productAddButtonId.setOnClickListener { addButtonClicked() }
        binding.openExcelButtonId.setOnClickListener { openExcel() }
    }

    private fun openExcel(){
        val filename = "razhodi.xls"
        val file = File(getExternalFilesDir(null),  filename)
        if(file.exists())
        {
            val uri = FileProvider.getUriForFile(applicationContext, applicationContext.packageName + ".provider", file)
            val intent = Intent(Intent.ACTION_VIEW)
            intent.data = uri
            intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
            startActivity(intent)
        }else{
            Log.d("Losho", "file does not exists!!!")
        }
    }

    private fun addButtonClicked() {
        val product = binding.productNameId.text
        val price = binding.productPriceId.text
        val count = binding.productCountId.text
        val productName = product.toString()
        val productPrice: Double? = price.toString().toDoubleOrNull()

        if (productName != "" && productPrice != null) {
            binding.productList.text = getString(R.string.item_output_text, productName, productPrice)

            val excel = ExcelFunctions(this)
            val filename = "razhodi.xls"
            val item = Item(productName, productPrice)

            val itemCount: Int = count.toString().toIntOrNull() ?: 1
            excel.writeToFile(filename, item, itemCount)
        }
        product.clear()
        price.clear()
    }
}