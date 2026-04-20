using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Group4337
{
    public class Client
    {
        [Key]
        public int Id { get; set; }

        [Required]
        [MaxLength(20)]
        public string ClientCode { get; set; } = string.Empty; // Код клиента

        [Required]
        [MaxLength(150)]
        public string FullName { get; set; } = string.Empty; // ФИО

        [Required]
        public DateTime BirthDate { get; set; } // Дата рождения

        [MaxLength(10)]
        public string PostalCode { get; set; } = string.Empty; // Индекс

        [MaxLength(100)]
        public string City { get; set; } = string.Empty; // Город

        [MaxLength(100)]
        public string Street { get; set; } = string.Empty; // Улица

        [MaxLength(10)]
        public string House { get; set; } = string.Empty; // Дом

        [MaxLength(10)]
        public string Apartment { get; set; } = string.Empty; // Квартира

        [MaxLength(200)]
        public string Email { get; set; } = string.Empty; // E-mail

        /// <summary>
        /// Возраст вычисляется программным способом из даты рождения
        /// </summary>
        [NotMapped]
        public int Age
        {
            get
            {
                var today = DateTime.Today;
                int age = today.Year - BirthDate.Year;
                if (BirthDate.Date > today.AddYears(-age))
                    age--;
                return age;
            }
        }

        /// <summary>
        /// Категория по возрасту
        /// </summary>
        [NotMapped]
        public string AgeCategory
        {
            get
            {
                if (Age >= 20 && Age <= 29)
                    return "Категория 1 (20–29)";
                else if (Age >= 30 && Age <= 39)
                    return "Категория 2 (30–39)";
                else if (Age >= 40)
                    return "Категория 3 (40+)";
                else
                    return "Без категории";
            }
        }
    }
}
