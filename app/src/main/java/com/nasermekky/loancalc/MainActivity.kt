
package com.nasermekky.loancalc

import android.content.Intent
import android.os.Bundle
import android.os.Environment
import android.widget.ArrayAdapter
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.recyclerview.widget.LinearLayoutManager
import com.nasermekky.loancalc.databinding.ActivityMainBinding
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import kotlin.math.pow

class MainActivity : AppCompatActivity() {
    private lateinit var binding: ActivityMainBinding
    private var schedule: List<ScheduleRow> = emptyList()
    private var resultText: String = ""

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        val methods = listOf(getString(R.string.flat), getString(R.string.reducing))
        val adapter = ArrayAdapter(this, android.R.layout.simple_spinner_dropdown_item, methods)
        binding.spinnerMethod.adapter = adapter

        binding.recyclerSchedule.layoutManager = LinearLayoutManager(this)

        binding.btnCalculate.setOnClickListener { calculate() }
        binding.btnClear.setOnClickListener { clearFields() }
        binding.btnExportCsv.setOnClickListener { exportCsv() }
        binding.btnExportXlsx.setOnClickListener { exportXlsx() }
        binding.btnShare.setOnClickListener { shareResults() }
    }

    private fun clearFields() {
        binding.inputPrincipal.setText("")
        binding.inputAnnualRate.setText("")
        binding.inputTermMonths.setText("")
        binding.txtResult.text = getString(R.string.result)
        binding.spinnerMethod.setSelection(0)
        schedule = emptyList()
        binding.recyclerSchedule.adapter = ScheduleAdapter(schedule)
    }

    private fun calculate() {
        val P = binding.inputPrincipal.text?.toString()?.toDoubleOrNull() ?: 0.0
        val annualRate = binding.inputAnnualRate.text?.toString()?.toDoubleOrNull()?.div(100.0) ?: 0.0
        val n = binding.inputTermMonths.text?.toString()?.toIntOrNull() ?: 0
        if (P <= 0 || annualRate < 0 || n <= 0) {
            toast("قيم غير صالحة")
            return
        }
        val r = annualRate / 12.0
        val methodIndex = binding.spinnerMethod.selectedItemPosition

        if (methodIndex == 0) {
            val years = n / 12.0
            val totalInterest = P * annualRate * years
            val total = P + totalInterest
            val installment = total / n
            resultText = "ثابتة: قسط ${"%,.2f".format(installment)} ، إجمالي فايدة ${"%,.2f".format(totalInterest)}"
            binding.txtResult.text = resultText
            schedule = emptyList()
        } else {
            val installment = if (r == 0.0) P / n else (P * r) / (1 - (1 + r).pow(-n))
            val total = installment * n
            val totalInterest = total - P
            resultText = "متناقصة: قسط ${"%,.2f".format(installment)} ، إجمالي فايدة ${"%,.2f".format(totalInterest)}"
            binding.txtResult.text = resultText
            schedule = generateSchedule(P, r, n, installment)
        }
        binding.recyclerSchedule.adapter = ScheduleAdapter(schedule)
    }

    private fun generateSchedule(P: Double, r: Double, n: Int, installment: Double): List<ScheduleRow> {
        var balance = P
        val list = mutableListOf<ScheduleRow>()
        for (i in 1..n) {
            val interest = balance * r
            val principal = installment - interest
            balance -= principal
            list.add(ScheduleRow(i, installment, principal, interest, if (balance < 0) 0.0 else balance))
        }
        return list
    }

    private fun exportCsv() {
        if (schedule.isEmpty()) {
            toast("لا يوجد جدول لتصديره")
            return
        }
        try {
            val dir = getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS)
            val file = File(dir, "schedule.csv")
            val fos = FileOutputStream(file)
            fos.write("No,Payment,Principal,Interest,Balance\n".toByteArray())
            for (row in schedule) {
                fos.write("${row.no},${row.payment},${row.principal},${row.interest},${row.balance}\n".toByteArray())
            }
            fos.close()
            toast("تم حفظ الملف CSV: ${file.absolutePath}")
        } catch (e: Exception) {
            toast("خطأ: ${e.message}")
        }
    }

    private fun exportXlsx() {
        if (schedule.isEmpty()) {
            toast("لا يوجد جدول لتصديره")
            return
        }
        try {
            val workbook = XSSFWorkbook()
            val sheet = workbook.createSheet("Schedule")

            val header = sheet.createRow(0)
            header.createCell(0).setCellValue("No")
            header.createCell(1).setCellValue("Payment")
            header.createCell(2).setCellValue("Principal")
            header.createCell(3).setCellValue("Interest")
            header.createCell(4).setCellValue("Balance")

            for ((i, row) in schedule.withIndex()) {
                val r = sheet.createRow(i + 1)
                r.createCell(0).setCellValue(row.no.toDouble())
                r.createCell(1).setCellValue(row.payment)
                r.createCell(2).setCellValue(row.principal)
                r.createCell(3).setCellValue(row.interest)
                r.createCell(4).setCellValue(row.balance)
            }

            val dir = getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS)
            val file = File(dir, "schedule.xlsx")
            val fos = FileOutputStream(file)
            workbook.write(fos)
            fos.close()
            workbook.close()
            toast("تم حفظ الملف Excel: ${file.absolutePath}")
        } catch (e: Exception) {
            toast("خطأ Excel: ${e.message}")
        }
    }

    private fun shareResults() {
        val sendIntent: Intent = Intent().apply {
            action = Intent.ACTION_SEND
            putExtra(Intent.EXTRA_TEXT, resultText)
            type = "text/plain"
        }
        startActivity(Intent.createChooser(sendIntent, "مشاركة النتائج"))
    }

    private fun toast(msg: String) {
        Toast.makeText(this, msg, Toast.LENGTH_SHORT).show()
    }
}
