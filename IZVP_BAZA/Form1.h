#pragma once

namespace CppCLRWinformsProjekt {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace ADOX;
	using namespace Microsoft::Office::Interop::Access;
	using namespace System::Data::OleDb;

	/// <summary>
	/// Zusammenfassung fьr Form1
	/// </summary>
	public ref class Form1 : public System::Windows::Forms::Form
	{
	public:
		Form1(void)
		{
			InitializeComponent();
			//
			//TODO: Konstruktorcode hier hinzufьgen.
			//
		}

	protected:
		/// <summary>
		/// Verwendete Ressourcen bereinigen.
		/// </summary>
		~Form1()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^ button1;
	protected:

	private:
		/// <summary>
		/// Erforderliche Designervariable.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Erforderliche Methode fьr die Designerunterstьtzung.
		/// Der Inhalt der Methode darf nicht mit dem Code-Editor geдndert werden.
		/// </summary>
		void InitializeComponent(void)
		{
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->SuspendLayout();
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(330, 226);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(158, 50);
			this->button1->TabIndex = 0;
			this->button1->Text = L"Структура";
			this->button1->UseVisualStyleBackColor = true;
			// 
			// Form1
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(860, 517);
			this->Controls->Add(this->button1);
			this->Name = L"Form1";
			this->Text = L"Form1";
			this->Load += gcnew System::EventHandler(this, &Form1::Form1_Load);
			this->ResumeLayout(false);

		}
#pragma endregion
	private: System::Void Form1_Load(System::Object^ sender, System::EventArgs^ e) {
		ADOX::Catalog^ Каталог = gcnew ADOX::Catalog();
		try
		{
			Каталог->Create("Provider=Microsoft.Jet."+"OLEDB.4.0;Data Source=d:\\Нова_БД.mbd");
			MessageBox::Show("База даних d:\\Нова_БД.mbd успішно створена", "Створення нової БД MS Access",
				MessageBoxButtons::OK, MessageBoxIcon::Information);
		}
		catch (System::Runtime::InteropServices::COMException^ Ситуація)
		{
			MessageBox::Show(Ситуація->Message, "Створення нової БД MS Access", MessageBoxButtons::OK,MessageBoxIcon::Warning);
		}
		finally
		{
			Каталог = nullptr;
		}
	}
	};
}