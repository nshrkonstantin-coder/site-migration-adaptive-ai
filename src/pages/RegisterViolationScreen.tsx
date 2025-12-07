import { useState } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Textarea } from '@/components/ui/textarea';
import { Label } from '@/components/ui/label';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { useNavigate } from 'react-router-dom';
import Icon from '@/components/ui/icon';
import { format } from 'date-fns';
import { useToast } from '@/hooks/use-toast';

const RegisterViolationScreen = () => {
  const navigate = useNavigate();
  const { toast } = useToast();
  const [isSubmitting, setIsSubmitting] = useState(false);
  
  const [formData, setFormData] = useState({
    date: format(new Date(), 'yyyy-MM-dd'),
    shop: '',
    section: '',
    objectInspected: '',
    description: '',
    auditor: 'Константин Ншрк, Инженер по технике безопасности',
    category: '',
    conditionType: '',
    hazardFactor: '',
    note: '',
    actions: '',
    responsibleName: '',
    dueDate: '',
    photoBase64: ''
  });

  const [photoPreview, setPhotoPreview] = useState<string | null>(null);

  const categories = [
    'Опасные условия',
    'Опасные действия',
    'Позитивное наблюдение',
    'Несоответствие требованиям'
  ];

  const conditionTypes = [
    'Работа на высоте',
    'Электробезопасность',
    'Пожарная безопасность',
    'Работа с химикатами',
    'Эксплуатация оборудования',
    'СИЗ',
    'Другое'
  ];

  const handleInputChange = (field: string, value: string) => {
    setFormData(prev => ({ ...prev, [field]: value }));
  };

  const handlePhotoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (file.size > 5 * 1024 * 1024) {
      toast({
        title: 'Ошибка',
        description: 'Размер файла не должен превышать 5 МБ',
        variant: 'destructive'
      });
      return;
    }

    const reader = new FileReader();
    reader.onloadend = () => {
      const base64String = reader.result as string;
      setPhotoPreview(base64String);
      setFormData(prev => ({ ...prev, photoBase64: base64String }));
    };
    reader.readAsDataURL(file);
  };

  const removePhoto = () => {
    setPhotoPreview(null);
    setFormData(prev => ({ ...prev, photoBase64: '' }));
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    
    if (!formData.shop || !formData.section || !formData.objectInspected || 
        !formData.description || !formData.category || !formData.conditionType) {
      toast({
        title: 'Ошибка',
        description: 'Заполните все обязательные поля',
        variant: 'destructive'
      });
      return;
    }

    setIsSubmitting(true);

    try {
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      toast({
        title: 'Успешно',
        description: 'Нарушение зарегистрировано',
      });
      
      navigate('/my-violations');
    } catch (error) {
      toast({
        title: 'Ошибка',
        description: 'Не удалось зарегистрировать нарушение',
        variant: 'destructive'
      });
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-50 to-gray-100 p-4 md:p-8">
      <div className="max-w-4xl mx-auto">
        <Button 
          variant="ghost" 
          onClick={() => navigate('/')}
          className="mb-4"
        >
          <Icon name="ArrowLeft" size={20} className="mr-2" />
          Назад
        </Button>
        
        <Card className="border-none shadow-lg">
          <CardHeader>
            <CardTitle className="text-2xl flex items-center gap-2">
              <Icon name="AlertTriangle" size={24} />
              Регистрация нарушений
            </CardTitle>
          </CardHeader>
          <CardContent>
            <form onSubmit={handleSubmit} className="space-y-6">
              <div className="space-y-2">
                <Label htmlFor="date">Дата нарушения *</Label>
                <Input
                  id="date"
                  type="date"
                  value={formData.date}
                  onChange={(e) => handleInputChange('date', e.target.value)}
                  required
                />
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-2">
                  <Label htmlFor="shop">Подразделение *</Label>
                  <Input
                    id="shop"
                    placeholder="Цех №1"
                    value={formData.shop}
                    onChange={(e) => handleInputChange('shop', e.target.value)}
                    required
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="section">Участок *</Label>
                  <Input
                    id="section"
                    placeholder="Участок сборки"
                    value={formData.section}
                    onChange={(e) => handleInputChange('section', e.target.value)}
                    required
                  />
                </div>
              </div>

              <div className="space-y-2">
                <Label htmlFor="objectInspected">Проверяемый объект *</Label>
                <Input
                  id="objectInspected"
                  placeholder="Станок №5"
                  value={formData.objectInspected}
                  onChange={(e) => handleInputChange('objectInspected', e.target.value)}
                  required
                />
              </div>

              <div className="space-y-2">
                <Label htmlFor="category">Категория наблюдения *</Label>
                <Select value={formData.category} onValueChange={(value) => handleInputChange('category', value)}>
                  <SelectTrigger>
                    <SelectValue placeholder="Выберите категорию" />
                  </SelectTrigger>
                  <SelectContent>
                    {categories.map((cat) => (
                      <SelectItem key={cat} value={cat}>{cat}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-2">
                <Label htmlFor="conditionType">Вид условий и действий *</Label>
                <Select value={formData.conditionType} onValueChange={(value) => handleInputChange('conditionType', value)}>
                  <SelectTrigger>
                    <SelectValue placeholder="Выберите вид" />
                  </SelectTrigger>
                  <SelectContent>
                    {conditionTypes.map((type) => (
                      <SelectItem key={type} value={type}>{type}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-2">
                <Label htmlFor="description">Описание наблюдения *</Label>
                <Textarea
                  id="description"
                  placeholder="Подробное описание нарушения..."
                  value={formData.description}
                  onChange={(e) => handleInputChange('description', e.target.value)}
                  rows={4}
                  required
                />
              </div>

              <div className="space-y-2">
                <Label htmlFor="hazardFactor">Опасные факторы</Label>
                <Textarea
                  id="hazardFactor"
                  placeholder="Описание опасных факторов"
                  value={formData.hazardFactor}
                  onChange={(e) => handleInputChange('hazardFactor', e.target.value)}
                  rows={3}
                />
              </div>

              <div className="space-y-2">
                <Label htmlFor="actions">Мероприятия</Label>
                <Textarea
                  id="actions"
                  placeholder="Рекомендуемые действия для устранения"
                  value={formData.actions}
                  onChange={(e) => handleInputChange('actions', e.target.value)}
                  rows={3}
                />
              </div>

              <div className="space-y-2">
                <Label htmlFor="auditor">Проверяющий *</Label>
                <Input
                  id="auditor"
                  value={formData.auditor}
                  onChange={(e) => handleInputChange('auditor', e.target.value)}
                  required
                />
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-2">
                  <Label htmlFor="responsibleName">Ответственный</Label>
                  <Input
                    id="responsibleName"
                    placeholder="ФИО ответственного"
                    value={formData.responsibleName}
                    onChange={(e) => handleInputChange('responsibleName', e.target.value)}
                  />
                </div>

                <div className="space-y-2">
                  <Label htmlFor="dueDate">Срок устранения</Label>
                  <Input
                    id="dueDate"
                    type="date"
                    value={formData.dueDate}
                    onChange={(e) => handleInputChange('dueDate', e.target.value)}
                  />
                </div>
              </div>

              <div className="space-y-2">
                <Label htmlFor="note">Примечание</Label>
                <Textarea
                  id="note"
                  placeholder="Дополнительная информация"
                  value={formData.note}
                  onChange={(e) => handleInputChange('note', e.target.value)}
                  rows={2}
                />
              </div>

              <div className="space-y-2">
                <Label>Фотография нарушения</Label>
                {!photoPreview ? (
                  <div className="border-2 border-dashed border-gray-300 rounded-lg p-8 text-center hover:border-primary transition-colors">
                    <input
                      type="file"
                      accept="image/*"
                      onChange={handlePhotoUpload}
                      className="hidden"
                      id="photo-upload"
                    />
                    <label htmlFor="photo-upload" className="cursor-pointer">
                      <Icon name="Camera" size={48} className="mx-auto mb-4 text-gray-400" />
                      <p className="text-sm text-muted-foreground">
                        Нажмите для загрузки фото
                      </p>
                      <p className="text-xs text-muted-foreground mt-1">
                        Максимальный размер: 5 МБ
                      </p>
                    </label>
                  </div>
                ) : (
                  <div className="relative">
                    <img 
                      src={photoPreview} 
                      alt="Превью" 
                      className="w-full h-64 object-cover rounded-lg"
                    />
                    <Button
                      type="button"
                      variant="destructive"
                      size="sm"
                      className="absolute top-2 right-2"
                      onClick={removePhoto}
                    >
                      <Icon name="X" size={16} />
                    </Button>
                  </div>
                )}
              </div>

              <div className="flex gap-4 pt-4">
                <Button
                  type="submit"
                  className="flex-1"
                  disabled={isSubmitting}
                >
                  {isSubmitting ? (
                    <>
                      <Icon name="Loader2" size={20} className="mr-2 animate-spin" />
                      Сохранение...
                    </>
                  ) : (
                    <>
                      <Icon name="Save" size={20} className="mr-2" />
                      Зарегистрировать нарушение
                    </>
                  )}
                </Button>
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => navigate('/')}
                  disabled={isSubmitting}
                >
                  Отмена
                </Button>
              </div>
            </form>
          </CardContent>
        </Card>
      </div>
    </div>
  );
};

export default RegisterViolationScreen;
