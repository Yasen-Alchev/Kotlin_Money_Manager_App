package com.example.nexttry

import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import com.example.nexttry.databinding.ActivityMainBinding

class MainActivity : AppCompatActivity() {
    lateinit var binding: ActivityMainBinding

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        binding.productAddButtonId.setOnClickListener { addButtonClicked() }
    }

    private fun addButtonClicked() {
        val product = binding.productNameId.text.toString()
        val price: Double? = binding.productPriceId.text.toString().toDoubleOrNull()

        if (product != "" && price != null) {
            binding.textView.text = getString(R.string.item_output_text, product, price)

            val excel = ExcelFunctions(this)
            val filename = "text.xls"
            val item = Item(product, price)
            excel.writeToFile(filename, item)
        }
    }
}