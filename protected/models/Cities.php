<?php

/**
 * This is the model class for table "cities".
 *
 * The followings are the available columns in table 'cities':
 * @property string $CityID
 * @property string $Name
 * @property string $Letter
 * @property string $Notes
 * @property integer $Enabled
 *
 * The followings are the available model relations:
 * @property PersonsCities[] $personsCities
 */
class Cities extends CActiveRecord
{
	/**
	 * @return string the associated database table name
	 */
	public function tableName()
	{
		return 'cities';
	}

	/**
	 * @return array validation rules for model attributes.
	 */
	public function rules()
	{
		// NOTE: you should only define rules for those attributes that
		// will receive user inputs.
		return array(
			array('Name, Enabled', 'required'),
			array('Enabled', 'numerical', 'integerOnly'=>true),
			array('Name', 'length', 'max'=>255),
			array('Letter', 'length', 'max'=>2),
			array('Notes', 'safe'),
			// The following rule is used by search().
			// @todo Please remove those attributes that should not be searched.
			array('CityID, Name, Letter, Notes, Enabled', 'safe', 'on'=>'search'),
		);
	}

	/**
	 * @return array relational rules.
	 */
	public function relations()
	{
		// NOTE: you may need to adjust the relation name and the related
		// class name for the relations automatically generated below.
		return array(
			'personsCities' => array(self::HAS_MANY, 'PersonsCities', 'CityID'),
		);
	}

	/**
	 * @return array customized attribute labels (name=>label)
	 */
	public function attributeLabels()
	{
		return array(
			'CityID' => 'City',
			'Name' => 'Name',
			'Letter' => 'Letter',
			'Notes' => 'Notes',
			'Enabled' => 'Enabled',
		);
	}

	/**
	 * Retrieves a list of models based on the current search/filter conditions.
	 *
	 * Typical usecase:
	 * - Initialize the model fields with values from filter form.
	 * - Execute this method to get CActiveDataProvider instance which will filter
	 * models according to data in model fields.
	 * - Pass data provider to CGridView, CListView or any similar widget.
	 *
	 * @return CActiveDataProvider the data provider that can return the models
	 * based on the search/filter conditions.
	 */
	public function search()
	{
		// @todo Please modify the following code to remove attributes that should not be searched.

		$criteria=new CDbCriteria;

		$criteria->compare('CityID',$this->CityID,true);
		$criteria->compare('Name',$this->Name,true);
		$criteria->compare('Letter',$this->Letter,true);
		$criteria->compare('Notes',$this->Notes,true);
		$criteria->compare('Enabled',$this->Enabled);

		return new CActiveDataProvider($this, array(
			'criteria'=>$criteria,
		));
	}

	/**
	 * Returns the static model of the specified AR class.
	 * Please note that you should have this exact method in all your CActiveRecord descendants!
	 * @param string $className active record class name.
	 * @return Cities the static model class
	 */
	public static function model($className=__CLASS__)
	{
		return parent::model($className);
	}
}
